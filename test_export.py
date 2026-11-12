"""Regression tests for file preservation and lossless PDF restoration."""
from io import BytesIO
from pathlib import Path
from tempfile import TemporaryDirectory
from unittest import TestCase, main
from unittest.mock import patch
from zipfile import ZipFile
import random

from PIL import Image
from pypdf import PdfWriter, PdfReader
from pypdf.generic import DictionaryObject, NameObject, NumberObject, DecodedStreamObject, IndirectObject
from precise_pptx import restore_and_verify, main as convert, encoded_image, sha, placement_match, image_placements


def fixture(root, im=None):
    if im is None:
        im = Image.new('RGB', (128, 80))
        im.putdata([(x % 256, (x * 3) % 256, (x * 11) % 256) for x in range(128 * 80)])
    png = BytesIO(); im.save(png, format='PNG')
    pptx = root / '含 空格.pptx'
    with ZipFile(pptx, 'w') as z:
        z.writestr('ppt/presentation.xml', '''<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:sldIdLst><p:sldId r:id="rId1"/></p:sldIdLst><p:sldSz cx="6400800" cy="2781300"/></p:presentation>''')
        z.writestr('ppt/_rels/presentation.xml.rels', '<Relationships><Relationship Id="rId1" Target="slides/slide1.xml"/></Relationships>')
        z.writestr('ppt/slides/slide1.xml', '''<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:pic><a:blip r:embed="rId1"/></p:pic></p:sld>''')
        z.writestr('ppt/slides/_rels/slide1.xml.rels', '<Relationships><Relationship Id="rId1" Target="../media/image1.png"/></Relationships>')
        z.writestr('ppt/media/image1.png', png.getvalue())
    writer = PdfWriter()
    page = writer.add_blank_page(width=504, height=219)
    img = DecodedStreamObject()
    img.set_data(im.tobytes())
    for name, value in {'/Type': '/XObject', '/Subtype': '/Image', '/ColorSpace': '/DeviceRGB'}.items():
        img[NameObject(name)] = NameObject(value)
    for name, value in {'/Width':im.width,'/Height':im.height,'/BitsPerComponent':8}.items():
        img[NameObject(name)] = NumberObject(value)
    ref = writer._add_object(img.flate_encode())
    page[NameObject('/Resources')] = DictionaryObject({NameObject('/XObject'): DictionaryObject({NameObject('/Im0'):ref})})
    content = DecodedStreamObject(); content.set_data(b'q 504 0 0 219 0 0 cm /Im0 Do Q')
    page[NameObject('/Contents')] = writer._add_object(content)
    pdf = root / 'native.pdf'; writer.write(pdf)
    return pptx, pdf


class ExportTests(TestCase):
    def test_compact_pdf_verifies_its_encoded_pixels_without_claiming_lossless(self):
        with TemporaryDirectory() as temp:
            im=Image.frombytes('RGB',(400,300),random.Random(123).randbytes(400*300*3))
            root=Path(temp);pptx,pdf=fixture(root,im)
            report=restore_and_verify(pptx,pdf,root/'result.pdf',strict=True,image_mode='compact')
            self.assertTrue(report['all_supported_source_images_matched'])
            self.assertFalse(report['all_supported_source_pixels_verified'])
            self.assertTrue(report['native_drawing_commands_unchanged'])
            self.assertEqual(report['restored_images'][0]['codec'],'jpeg-q95-444')
            self.assertEqual(report['restored_images'][0]['size'],[400,300])

    def test_unreferenced_old_images_are_removed(self):
        with TemporaryDirectory() as temp:
            root=Path(temp);pptx,pdf=fixture(root)
            writer=PdfWriter(clone_from=pdf)
            duplicate=DecodedStreamObject()
            original=writer.pages[0].images[0].indirect_reference.get_object()
            duplicate.update(original);duplicate.set_data(b'unused image')
            writer._add_object(duplicate)
            writer.write(pdf)
            result=root/'result.pdf'
            restore_and_verify(pptx,pdf,result,strict=True)
            verified=PdfReader(result);count=0
            for gen,objects in verified.xref.items():
                for num in objects:
                    if num:
                        obj=verified.get_object(IndirectObject(num,gen,verified))
                        if hasattr(obj,'get') and obj.get('/Subtype')=='/Image':count+=1
            self.assertEqual(count,1)

    def test_compact_codec_preserves_resolution_and_reports_loss(self):
        # Smooth color photograph-like data has enough colors to use JPEG.
        im=Image.frombytes('RGB',(400,300),random.Random(123).randbytes(400*300*3))
        stream,decoded_hash,codec=encoded_image(im,'compact')
        self.assertEqual(codec,'jpeg-q95-444')
        decoded=Image.open(BytesIO(stream._data)).convert('RGB')
        self.assertEqual(decoded.size,im.size)
        self.assertEqual(sha(decoded.tobytes()),decoded_hash)
        self.assertNotEqual(decoded_hash,sha(im.tobytes()))
        self.assertLess(len(stream._data),len(im.tobytes()))

    def test_simple_graphics_remain_lossless_in_compact_mode(self):
        im=Image.new('RGB',(600,400),'white')
        stream,decoded_hash,codec=encoded_image(im,'compact')
        self.assertEqual(codec,'png-predictor-lossless')
        self.assertEqual(decoded_hash,sha(im.tobytes()))

    def test_exact_source_size_replaces_native_rounding(self):
        with TemporaryDirectory() as temp:
            root=Path(temp);pptx,pdf=fixture(root)
            with ZipFile(pptx) as z:parts={n:z.read(n) for n in z.namelist()}
            parts['ppt/presentation.xml']=parts['ppt/presentation.xml'].replace(b'cx="6400800"',b'cx="6395720"')
            with ZipFile(pptx,'w') as z:
                for n,v in parts.items():z.writestr(n,v)
            report=restore_and_verify(pptx,pdf,root/'result.pdf',strict=True)
            self.assertEqual(report['width_pt'],503.6)
            self.assertTrue(report['native_drawing_commands_unchanged'])

    def test_placement_requires_unique_geometry_and_similar_picture(self):
        im=Image.new('RGB',(100,80),'red');box=[10,20,50,40]
        src={'image':im,'thumbnail':im.resize((64,40)),'placements':[(0,box)]}
        self.assertEqual(placement_match(im,[box],{'a':src},0)[1],'a')
        self.assertIsNone(placement_match(im,[box],{'a':src,'b':src},0))
        self.assertIsNone(placement_match(Image.new('RGB',(100,80),'blue'),[box],{'a':src},0))
        self.assertIsNone(placement_match(im,[[0,0,50,40]],{'a':src},0))

    def test_native_placement_tracks_translation_and_scale(self):
        with TemporaryDirectory() as temp:
            _,pdf=fixture(Path(temp));p=PdfReader(pdf).pages[0]
            boxes=list(image_placements(p).values())
            self.assertEqual(boxes,[[[0.,0.,504.,219.]]])

    def test_full_image_pixels_and_vector_commands_survive(self):
        with TemporaryDirectory() as temp:
            root = Path(temp); pptx, pdf = fixture(root)
            report = restore_and_verify(pptx, pdf, root/'result.pdf', strict=True)
            self.assertEqual(len(report['restored_images']), 1)
            self.assertTrue(report['all_supported_source_pixels_verified'])
            self.assertTrue(report['nested_vector_streams_unchanged'])

    def test_wrong_page_size_is_rejected(self):
        with TemporaryDirectory() as temp:
            root = Path(temp); pptx, pdf = fixture(root)
            writer = PdfWriter(); writer.add_blank_page(width=400,height=300); writer.write(pdf)
            with self.assertRaisesRegex(ValueError, 'dimensions'):
                restore_and_verify(pptx, pdf, root/'result.pdf')
            self.assertFalse((root/'result.pdf').exists())

    def test_source_change_keeps_existing_outputs(self):
        with TemporaryDirectory() as temp:
            root = Path(temp); pptx, pdf = fixture(root)
            out = root/'out'; out.mkdir()
            target = out/pptx.with_suffix('.pdf').name; target.write_bytes(b'previous PDF')
            def changing_export(snapshot, native, original, timeout):
                native.write_bytes(pdf.read_bytes())
                original.write_bytes(b'changed by user')
            with patch('precise_pptx.native_export', changing_export):
                with self.assertRaisesRegex(RuntimeError, 'Source changed'):
                    convert([str(pptx),'--output',str(out),'--overwrite','--copy-pptx'])
            self.assertEqual(target.read_bytes(), b'previous PDF')
            self.assertFalse((out/pptx.name).exists())

    def test_duplicate_names_fail_before_export(self):
        with TemporaryDirectory() as temp:
            root=Path(temp); pptx, _ = fixture(root)
            other=root/'other'; other.mkdir(); (other/pptx.name).write_bytes(pptx.read_bytes())
            with patch('precise_pptx.native_export') as export:
                with self.assertRaises(SystemExit):
                    convert([str(pptx),str(other/pptx.name),'--output',str(root/'out')])
                export.assert_not_called()

if __name__ == '__main__': main()
