# !/usr/bin/env python3
# -*- coding: utf-8 -*-
# @Author 白孟阳
# 单元测试
import unittest
from RunbowTransfer import checkArea

class testCheckArea(unittest.TestCase):
    def setUp(self):
        print('setUp...')

    def tearDown(self):
        print('tearDown...')
        
    def test_checkArea(self):
        area,location=checkArea("C-14-3-04")
        self.assertEqual(area, "AU1")
        self.assertEqual(location, "C-14-3-04")
        
if __name__ == '__main__':
    unittest.main()