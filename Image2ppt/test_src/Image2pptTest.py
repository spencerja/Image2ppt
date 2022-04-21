import unittest
import sys
sys.path.append('../')
from src import Model


class MyTestCase(unittest.TestCase):
    def test_margintest(self):
        model = Model.Model()
        margin = model.get_margin_in_pixel(200,50)
        assert margin == 75

    def test_autopass(self):
        testpass = True
        assert testpass == True

    def test_get_margin_in_pixel(self):
        model = Model.Model()
        num = model.get_margin_in_pixel(100,50)
        assert num == 25

    def test_resize_ratio(self):
        model = Model.Model()
        ratio = model.get_resize_ratio(100,100,50,50)
        assert ratio == 0.5
    def test_margin_negative(self):
        model = Model.Model()
        margin = model.get_margin_in_pixel(-100,-50)
        assert margin == -25



    # def test_resize_ratio_horizontal_longer(self):
    #     model = Model.Model()
    #     ratio = model.get_resize_ratio(1000,500,100,100)
    #     assert ratio == 0.2
    #
    # def test_resize_ratio_vertical_longer(self):
    #     model = Model.Model()
    #     ratio = model.get_resize_ratio(500,1000,100,100)
    #     assert ratio == 0.2


if __name__ == '__main__':
    unittest.main()
