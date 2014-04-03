# -*- coding: utf-8 -*-

import xlrd
from xlrd import *
import json
from collections import OrderedDict


class FileNotInformed(Exception):
	pass


class EmptyPage(Exception):
	pass


class BaseXLS(object):

	def __init__(self, filepath=None, only_fields=None, limit=None):
		self.only_fields = only_fields
		self.limit = limit
		if not filepath:
			raise FileNotInformed('The XLS file was not informed.')
		self.open_xls(filepath)

	def open_xls(self, filepath):
		"""
		Just open the xls file
		"""
		self.xls_file_path = filepath
		self.xls_file = xlrd.open_workbook(filepath)
		return self

	def parse_row(self, row):
		"""
		Parse the row and returns a list with the values
		"""
		output = []
		for cell in row:
			if cell.ctype == 0:
				output.append(None)
			elif cell.ctype == 1:
				output.append(cell.value)
			elif cell.ctype == 2 or cell.ctype == 3:
				output.append(float(cell.value))
			elif cell.ctype == 4:
				output.append(bool(cell.value))
			elif cell.ctype == 5:
				output.append(u'ERROR')
		return output


	def read_xls(self, page=0):
		"""
		Read the XLS and returns a list of LISTS
		"""
		# Shure you can pass a alternative file.
		xls_data = self.xls_file.sheet_by_index(page)
		if xls_data.nrows <= 0:
			raise EmptyPage('The page %d has no values.' % page)
		total_lines = self.limit if self.limit and self.limit <= xls_data.nrows else xls_data.nrows 
		output_data = []
		for row_number in xrange(0, total_lines):
			row = self.parse_row(row=xls_data.row(row_number))
			output_data.append(row)
		return output_data


class FromXLS(BaseXLS):

	def to_python(self, ordered=False):
		""" 
		Get every line of the xls file
		and then create a dictionary using the header
		"""
		if ordered:
			dict_maker = OrderedDict
		else:
			dict_maker = dict

		xls_data = self.read_xls()
		output_json = []
		# The header is always the first line.
		header = xls_data[0]
		for row in xls_data[1:]:
			data = dict_maker(zip(header, row))
			output_json.append(data)
		return output_json

	def to_json(self):
		return json.dumps(self.to_python())
