#!/usr/bin/env python3
"""PowerPoint-native PDF export with verified, lossless source-image restoration."""
from __future__ import annotations
import argparse
from contextlib import contextmanager
from datetime import datetime
import fcntl
import hashlib
from io import BytesIO
import json
import logging
import os
from pathlib import Path
import posixpath
import shutil
import subprocess
import sys
import tempfile
import xml.etree.ElementTree as ET
from zipfile import ZipFile

from PIL import Image, ImageChops, ImageStat
from pypdf import PdfReader, PdfWriter
from pypdf.generic import NameObject, NumberObject

ROOT = Path(__file__).resolve().parent
VERSION = "1.1.0"
NS = {"p": "http://schemas.openxmlformats.org/presentationml/2006/main",
      "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
      "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
      "mc": "http://schemas.openxmlformats.org/markup-compatibility/2006"}
logging.getLogger("pypdf").setLevel(logging.ERROR)

def sha(data):
    return hashlib.sha256(data).hexdigest()

def source_info(path):
    """Read slide order and only raster pictures actually referenced by slides."""
    images = {}
    unsupported = []
    with ZipFile(path) as z:
        pres = ET.fromstring(z.read("ppt/presentation.xml"))
        size = pres.find("p:sldSz", NS)
        rels = {r.get("Id"): r.get("Target") for r in ET.fromstring(z.read("ppt/_rels/presentation.xml.rels"))}
        slides = []
        for sid in pres.find("p:sldIdLst", NS):
            target = rels[sid.get("{" + NS["r"] + "}id")]
            slides.append(posixpath.normpath(posixpath.join("ppt", target)).lstrip("/"))
        for page_index, slide in enumerate(slides):
            tree = ET.fromstring(z.read(slide))
            # Office Math fallback bitmaps are not displayed by modern PowerPoint.
            for ac in tree.findall(".//mc:AlternateContent", NS):
                fallback = ac.find("mc:Fallback", NS)
                if ac.find("mc:Choice", NS) is not None and fallback is not None:
                    ac.remove(fallback)
            relfile = posixpath.join(posixpath.dirname(slide), "_rels", posixpath.basename(slide) + ".rels")
            if relfile not in z.namelist():
                continue
            picrels = {r.get("Id"): r for r in ET.fromstring(z.read(relfile))}
            for pic in tree.findall(".//p:pic", NS):
                blip = pic.find(".//a:blip", NS)
                rid = blip.get("{" + NS["r"] + "}embed") if blip is not None else None
                rel = picrels.get(rid)
                if rel is None or rel.get("TargetMode") == "External":
                    unsupported.append("Linked picture: " + slide)
                    continue
                name = posixpath.normpath(posixpath.join(posixpath.dirname(slide), rel.get("Target"))).lstrip("/")
                try:
                    im = Image.open(BytesIO(z.read(name)))
                    im.load()
                except (OSError, KeyError):
                    # SVG/EMF/WMF stay in PowerPoint's native export.
                    continue
                if im.mode not in ("RGB", "L", "RGBA"):
                    unsupported.append("Unsupported color mode: " + name)
                    continue
                if im.mode == "RGBA" and im.getchannel("A").getextrema() != (255, 255):
                    unsupported.append("Transparent source picture: " + name)
                    continue
                im = im.convert("RGB")
                key = sha(im.tobytes() + str(im.size).encode())
                entry = images.setdefault(key, {"name": name, "image": im,
                    "thumbnail": im.resize((64, 40)), "pixel_sha256": sha(im.tobytes()), "pages": []})
                if page_index not in entry["pages"]:
                    entry["pages"].append(page_index)
        return {"slide_count": len(slides), "width_pt": int(size.get("cx"))/12700,
                "height_pt": int(size.get("cy"))/12700, "images": images, "unsupported": unsupported}

def image_match(image, sources):
    """Only accept a unique, very close full-image match; never guess a crop."""
    thumb = image.convert("RGB").resize((64, 40))
    aspect = image.width / image.height
    scores = []
    for key, source in sources.items():
        original = source["image"]
        if abs(aspect / (original.width / original.height) - 1) > .012:
            continue
        diff = ImageStat.Stat(ImageChops.difference(thumb, source["thumbnail"]))
        mse = sum(v*v for v in diff.rms)/3
        scores.append((mse, key))
    scores.sort()
    if not scores or scores[0][0] > 36:
        return None
    if len(scores) > 1 and scores[1][0] < max(64, scores[0][0]*4):
        return None
    return scores[0]

def font_status(reader):
    found = {}
    seen = set()
    def walk(resources):
        resources = resources.get_object()
        for ref in resources.get("/Font", {}).values():
            font = ref.get_object()
            descfont = font["/DescendantFonts"][0].get_object() if font.get("/DescendantFonts") else font
            desc = descfont.get("/FontDescriptor", {})
            if hasattr(desc, "get_object"):
                desc = desc.get_object()
            embedded = font.get("/Subtype") == "/Type3" or any(k in desc for k in ("/FontFile", "/FontFile2", "/FontFile3"))
            found[str(font.get("/BaseFont", "Type3"))] = embedded
        for ref in resources.get("/XObject", {}).values():
            identity = (getattr(ref, "idnum", id(ref)), getattr(ref, "generation", 0))
            if identity in seen:
                continue
            seen.add(identity)
            obj = ref.get_object()
            if obj.get("/Resources"):
                walk(obj["/Resources"])
    for page in reader.pages:
        walk(page["/Resources"])
    return found

def content_bytes(page):
    content = page.get_contents()
    return content.get_data() if content else b""

def drawing_streams(reader):
    """Include nested Form XObjects, where PowerPoint also stores vector drawing."""
    streams = []
    def walk(resources, prefix, ancestors):
        for name, ref in resources.get_object().get("/XObject", {}).items():
            obj = ref.get_object()
            if obj.get("/Subtype") != "/Form":
                continue
            identity = (getattr(ref, "idnum", id(obj)), getattr(ref, "generation", 0))
            if identity in ancestors:
                continue
            key = prefix + str(name)
            streams.append((key, sha(obj.get_data())))
            if obj.get("/Resources"):
                walk(obj["/Resources"], key, ancestors | {identity})
    for index, page in enumerate(reader.pages):
        streams.append((str(index), sha(content_bytes(page))))
        walk(page.get("/Resources", {}), str(index), set())
    return streams

def restore_and_verify(pptx, native_pdf, candidate, strict=False):
    info = source_info(pptx)
    reader = PdfReader(native_pdf)
    if reader.is_encrypted:
        raise ValueError("Encrypted PDF is not supported")
    if len(reader.pages) != info["slide_count"]:
        raise ValueError("PDF page count differs from PPTX slide count")
    for page in reader.pages:
        if abs(float(page.mediabox.width)-info["width_pt"]) > 1 or abs(float(page.mediabox.height)-info["height_pt"]) > 1:
            raise ValueError("PDF dimensions differ from the PowerPoint slide")
    writer = PdfWriter(clone_from=reader)
    replacements = []
    used = set()
    seen = set()
    for page_no, page in enumerate(writer.pages):
        for image in list(page.images):
            ref = image.indirect_reference
            if ref is None or (ref.idnum, ref.generation) in seen:
                continue
            seen.add((ref.idnum, ref.generation))
            obj = ref.get_object()
            if obj.get("/SMask") or obj.get("/Mask") or obj.get("/Decode"):
                continue
            page_sources = {k: v for k, v in info["images"].items() if page_no in v["pages"]}
            matched = image_match(image.image, page_sources)
            if matched is None:
                continue
            mse, key = matched
            original = info["images"][key]["image"]
            # Do not substitute a lower-resolution original for a larger native render.
            if original.width < image.image.width or original.height < image.image.height:
                continue
            used.add(key)
            # Populate pypdf's decoded stream cache before switching the codec.
            obj.get_data()
            obj[NameObject("/Filter")] = NameObject("/FlateDecode")
            obj[NameObject("/Width")] = NumberObject(original.width)
            obj[NameObject("/Height")] = NumberObject(original.height)
            obj[NameObject("/BitsPerComponent")] = NumberObject(8)
            # Native RGB color spaces/ICC profiles are retained when compatible.
            if image.image.mode != "RGB":
                obj[NameObject("/ColorSpace")] = NameObject("/DeviceRGB")
            obj.pop("/DecodeParms", None)
            obj.set_data(original.tobytes())
            replacements.append({"page": page_no+1, "image": image.name.split(".")[0],
                                 "source": info["images"][key]["name"], "size": list(original.size),
                                 "pixel_sha256": sha(original.tobytes()), "match_mse": round(mse, 6)})
    unmatched = [s["name"] for key, s in info["images"].items() if key not in used]
    if strict and (unmatched or info["unsupported"]):
        raise ValueError("Strict image check failed: " + "; ".join(unmatched + info["unsupported"]))
    writer.write(candidate)
    verified = PdfReader(candidate)
    if drawing_streams(reader) != drawing_streams(verified):
        raise AssertionError("Nested vector drawing commands changed")
    for old, new in zip(reader.pages, verified.pages):
        if content_bytes(old) != content_bytes(new) or list(old.mediabox) != list(new.mediabox):
            raise AssertionError("Page drawing commands or dimensions changed")
        if old.extract_text() != new.extract_text():
            raise AssertionError("PDF text changed")
    for record in replacements:
        page = verified.pages[record["page"]-1]
        image = next(im for im in page.images if im.name.split(".")[0] == record["image"])
        if list(image.image.size) != record["size"] or sha(image.image.convert("RGB").tobytes()) != record["pixel_sha256"]:
            raise AssertionError("Source pixels were not preserved")
    fonts = font_status(verified)
    if font_status(reader) != fonts:
        raise AssertionError("PDF font resources changed")
    return {"source_sha256": sha(Path(pptx).read_bytes()), "pdf_sha256": sha(Path(candidate).read_bytes()),
            "pages": len(verified.pages), "width_pt": float(verified.pages[0].mediabox.width),
            "height_pt": float(verified.pages[0].mediabox.height), "fonts_embedded": fonts,
            "native_drawing_commands_unchanged": True, "nested_vector_streams_unchanged": True,
            "restored_images": replacements,
            "unmatched_source_images": unmatched, "unsupported_source_images": info["unsupported"],
            "all_supported_source_pixels_verified": not unmatched and not info["unsupported"]}

def native_export(pptx, pdf, original, timeout):
    if sys.platform != "darwin":
        raise RuntimeError("PowerPoint export requires macOS")
    script = ROOT / "powerpoint_export.applescript"
    try:
        result = subprocess.run(["/usr/bin/osascript", str(script), str(pptx), str(pdf), str(original), str(max(10, timeout-5))],
                                capture_output=True, text=True, timeout=timeout)
    except subprocess.TimeoutExpired as error:
        raise RuntimeError("PowerPoint 导出超时。请查看 PowerPoint 是否有文件访问、修复或其他弹窗，处理后重试。源文件和已有输出未改动。") from error
    if result.returncode:
        if "-1712" in result.stderr:
            raise RuntimeError("PowerPoint 等待超时。请检查并处理 PowerPoint 弹窗后重试；首次使用需要允许访问 PPTX-Precise/Work 文件夹。")
        raise RuntimeError("PowerPoint export failed: " + result.stderr.strip())
    if not pdf.is_file() or pdf.stat().st_size < 100:
        raise RuntimeError("PowerPoint did not produce a PDF")

def publish(source, destination, overwrite, backup_dir):
    if destination.exists():
        if not overwrite:
            raise FileExistsError(str(destination) + " exists; use --overwrite")
        backup_dir.mkdir(parents=True, exist_ok=True)
        shutil.copy2(destination, backup_dir / destination.name)
    # Atomic replacement within the destination filesystem.
    with tempfile.NamedTemporaryFile(dir=destination.parent, prefix=".pptx-precise-", delete=False) as f:
        temporary = Path(f.name)
    try:
        shutil.copy2(source, temporary)
        os.replace(temporary, destination)
    finally:
        temporary.unlink(missing_ok=True)

@contextmanager
def export_workspace():
    """A stable folder lets PowerPoint remember its macOS access grant."""
    stage = Path.home() / "Library/Application Support/PPTX-Precise/Work"
    stage.mkdir(parents=True, exist_ok=True)
    fd, name = tempfile.mkstemp(prefix="PPTX-Precise-", suffix=".pptx", dir=stage)
    os.close(fd)
    snapshot = Path(name)
    native = snapshot.with_suffix(".native.pdf")
    candidate = snapshot.with_suffix(".verified.pdf")
    try:
        yield snapshot, native, candidate
    finally:
        for path in (snapshot, native, candidate):
            path.unlink(missing_ok=True)

def main(argv=None):
    parser = argparse.ArgumentParser(description="macOS PowerPoint 精准转 PDF：原生排版＋原图像素恢复＋校验")
    parser.add_argument("--version", action="version", version=VERSION)
    parser.add_argument("inputs", nargs="+", type=Path, help="PPTX files or folders")
    parser.add_argument("--output", type=Path, required=True)
    parser.add_argument("--copy-pptx", action="store_true", help="Copy the exact saved source alongside PDF")
    parser.add_argument("--overwrite", action="store_true", help="Back up and replace existing outputs")
    parser.add_argument("--strict-images", action="store_true", help="Fail if any referenced raster source could not be restored")
    parser.add_argument("--native-pdf-dir", type=Path, help="Use already exported matching native PDFs (diagnostic mode)")
    parser.add_argument("--timeout", type=int, default=300)
    args = parser.parse_args(argv)
    files = []
    for item in args.inputs:
        files.extend(sorted(p for p in item.iterdir() if p.suffix.lower() == ".pptx") if item.is_dir() else [item])
    files = list(dict.fromkeys(p.resolve() for p in files if not p.name.startswith("~$")))
    if not files or any(p.suffix.lower() != ".pptx" or not p.is_file() for p in files):
        parser.error("Select existing PPTX files")
    if len({p.stem.casefold() for p in files}) != len(files):
        parser.error("Input names collide in the output folder")
    out = args.output.resolve()
    out.mkdir(parents=True, exist_ok=True)
    for path in files:
        for target in [out / (path.stem + ".pdf")] + ([out / path.name] if args.copy_pptx and path.parent != out else []):
            if target.exists() and not args.overwrite:
                parser.error(str(target) + " exists; pass --overwrite")
    stamp = datetime.now().strftime("%Y%m%d-%H%M%S-%f")
    reports = out / ".pptx-precise-reports" / stamp
    backups = out / ".pptx-precise-backups" / stamp
    reports.mkdir(parents=True, exist_ok=True)
    lockpath = Path(tempfile.gettempdir()) / ("pptx-precise-"+str(os.getuid())+".lock")
    fd = os.open(lockpath, os.O_CREAT | os.O_RDWR | getattr(os, "O_NOFOLLOW", 0), 0o600)
    with os.fdopen(fd, "w") as lock:
        try:
            fcntl.flock(lock, fcntl.LOCK_EX | fcntl.LOCK_NB)
        except BlockingIOError:
            raise RuntimeError("Another PPTX Precise conversion is already running")
        for index, source in enumerate(files, 1):
            print(f"[{index}/{len(files)}] {source.name}", flush=True)
            source_bytes = source.read_bytes()
            initial_hash = sha(source_bytes)
            with export_workspace() as (snapshot, native, candidate):
                snapshot.write_bytes(source_bytes)
                if args.native_pdf_dir:
                    shutil.copy2(args.native_pdf_dir / (source.stem+".pdf"), native)
                else:
                    native_export(snapshot, native, source, args.timeout)
                record = restore_and_verify(snapshot, native, candidate, args.strict_images)
                if initial_hash != sha(source.read_bytes()):
                    raise RuntimeError("Source changed during conversion. Save it and rerun: " + str(source))
                record["source"] = str(source)
                record["tool_version"] = VERSION
                record["exported_at"] = datetime.now().astimezone().isoformat()
                record["engine"] = "external-native-pdf" if args.native_pdf_dir else "Microsoft PowerPoint for macOS"
                record["native_pdf_sha256"] = sha(native.read_bytes())
                if any(not v for v in record["fonts_embedded"].values()):
                    print("  Warning: some native PDF fonts are not embedded", flush=True)
                if not record["all_supported_source_pixels_verified"]:
                    print("  Warning: some image effects were retained from the native export; see report", flush=True)
                publish(candidate, out / (source.stem+".pdf"), args.overwrite, backups)
                if args.copy_pptx and source.parent != out:
                    publish(snapshot, out / source.name, args.overwrite, backups)
                (reports / (source.stem+".json")).write_text(json.dumps(record, ensure_ascii=False, indent=2))
                print(f"  OK: {record['pages']} page(s), {len(record['restored_images'])} original images verified", flush=True)
    print("Saved to: " + str(out), flush=True)
    return 0

if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except (OSError, ValueError, RuntimeError, AssertionError, subprocess.SubprocessError) as error:
        print("ERROR: " + str(error), file=sys.stderr)
        raise SystemExit(1)
