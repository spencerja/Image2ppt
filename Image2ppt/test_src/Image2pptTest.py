import unittest
import sys
sys.path.append('../')
from src import Model


class MyTestCase(unittest.TestCase):
    # def test_get_resize_ratio_horizontal_image_vertical_slide_panel(self):
    #     ap_slide = Image2ppt.AppendSlide("")
    #     result = ap_slide.get_resize_ratio(1920,1080,240,270)
    #     self.assertEqual(0.125,result)
    #
    # def test_get_resize_ratio_vertical_image_vertical_slide_panel(self):
    #     ap_slide = Image2ppt.AppendSlide("")
    #     result = ap_slide.get_resize_ratio(480,1080,240,270)
    #     self.assertEqual(0.25,result)
    def test_get_margin_in_pixel(self):
        model = Model.Model()
        num = model.get_margin_in_pixel(100,50)
        assert num == 25

if __name__ == '__main__':
    unittest.main()
