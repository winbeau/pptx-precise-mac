"""Regression tests for file preservation and lossless PDF restoration."""
from io import BytesIO
from pathlib import Path
from tempfile import TemporaryDirectory
from unittest import TestCase, main
from unittest.mock import patch
from zipfile import ZipFile

from PIL import Image
from pypdf import PdfWriter
from pypdf.generic import DictionaryObject, NameObject, NumberObject, DecodedStreamObject
from precise_pptx import restore_and_verify, main as convert


def fixture(root):
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
    for name, value in {'/Width':128,'/Height':80,'/BitsPerComponent':8}.items():
        img[NameObject(name)] = NumberObject(value)
    ref = writer._add_object(img.flate_encode())
    page[NameObject('/Resources')] = DictionaryObject({NameObject('/XObject'): DictionaryObject({NameObject('/Im0'):ref})})
    content = DecodedStreamObject(); content.set_data(b'q 504 0 0 219 0 0 cm /Im0 Do Q')
    page[NameObject('/Contents')] = writer._add_object(content)
    pdf = root / 'native.pdf'; writer.write(pdf)
    return pptx, pdf


class ExportTests(TestCase):
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
