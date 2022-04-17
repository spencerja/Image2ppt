import unittest
import sys
sys.path.append('../')
from src import Model


class MyTestCase(unittest.TestCase):

    def test_get_margin_in_pixel(self):
        model = Model.Model()
        num = model.get_margin_in_pixel(100,50)
        assert num == 25

    def test_resize_ratio(self):
        model = Model.Model()
        ratio = model.get_resize_ratio(100,100,50,50)
        assert ratio == 0.5

if __name__ == '__main__':
    unittest.main()
