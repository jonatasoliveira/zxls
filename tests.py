#-*- coding: utf-8 -*-
import unittest
from zxls import *


class FromXLSTest(unittest.TestCase):

	def test_invalid_filename(self):
		with self.assertRaises(IOError):
			xls = FromXLS('this_is_not_a_valid_path.xls')

	def test_not_xls_file(self):
		with self.assertRaises(XLRDError):
			xls = FromXLS('__init__.py')

	def test_header_xls_ordered(self):
		header = ['Band', 'Genre', 'Best Album']
		xls = FromXLS('xls/teste.xls').to_python()[0]
		self.assertNotEqual(xls.keys(), header)
		xls = FromXLS('xls/teste.xls').to_python(ordered=True)[0]
		self.assertEqual(xls.keys(), header)

	def test_empty_page(self):
		with self.assertRaises(EmptyPage):
			xls = FromXLS('xls/empty.xls').to_python()


if __name__ == '__main__':
    unittest.main()
