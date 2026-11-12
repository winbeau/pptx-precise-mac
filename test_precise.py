import tempfile
from pathlib import Path
import unittest
from PIL import Image, ImageDraw
from precise_pptx import image_match, publish, sha

def sample(color):
    im=Image.new('RGB',(160,100),color)
    ImageDraw.Draw(im).rectangle((15,25,75,80),fill='white')
    return im

def source(im):
    return {'image':im,'thumbnail':im.resize((64,40))}

class PrecisionTests(unittest.TestCase):
    def test_unique_reduced_image_matches_source(self):
        red,blue=sample('red'),sample('blue')
        self.assertEqual(image_match(red.resize((80,50)),{'red':source(red),'blue':source(blue)})[1],'red')

    def test_ambiguous_sources_are_not_guessed(self):
        red=sample('red')
        self.assertIsNone(image_match(red,{'a':source(red),'b':source(red.copy())}))

    def test_crop_is_not_replaced_with_an_unrelated_full_frame(self):
        im=sample('red')
        self.assertIsNone(image_match(im.crop((0,0,80,100)),{'a':source(im)}))

    def test_publish_preserves_old_output_in_backup(self):
        with tempfile.TemporaryDirectory() as td:
            d=Path(td);src=d/'source';dest=d/'target';backup=d/'backup'
            src.write_bytes(b'new');dest.write_bytes(b'old')
            with self.assertRaises(FileExistsError):publish(src,dest,False,backup)
            self.assertEqual(dest.read_bytes(),b'old')
            publish(src,dest,True,backup)
            self.assertEqual(dest.read_bytes(),b'new')
            self.assertEqual((backup/'target').read_bytes(),b'old')

if __name__=='__main__':unittest.main()
