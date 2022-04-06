import unittest
from Image2ppt.src import Image2ppt


class MyTestCase(unittest.TestCase):
    def test_get_resize_ratio_horizontal_image_vertical_slide_panel(self):
        ap_slide = Image2ppt.AppendSlide("")
        result = ap_slide.get_resize_ratio(1920,1080,240,270)
        self.assertEqual(0.125,result)

    def test_get_resize_ratio_vertical_image_vertical_slide_panel(self):
        ap_slide = Image2ppt.AppendSlide("")
        result = ap_slide.get_resize_ratio(480,1080,240,270)
        self.assertEqual(0.25,result)



if __name__ == '__main__':
    unittest.main()
