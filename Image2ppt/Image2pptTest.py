import unittest
import Image2ppt

class MyTestCase(unittest.TestCase):
    def test_addition(self):
        add = Image2ppt.AddTest()
        result = add.add_test(1,2)
        self.assertEqual(3, result)  # add assertion here


if __name__ == '__main__':
    unittest.main()
