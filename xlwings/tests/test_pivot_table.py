import sys
sys.path.append(r'/Users/Harrison/anaconda3/envs/XLWingsDev/xlwings/')
import xlwings as xw
from xlwings.tests.common import TestBase, this_dir
import unittest

test_sheet = 'test_pivot.xlsx'
SPEC = '/Applications/Microsoft Office 2011/Microsoft Excel.app'


class test_pivot_table(unittest.TestCase):

	@classmethod
	def setUpClass(cls):
		cls.app1 = xw.App(visible=False, spec=SPEC)

	def setUp(self):
		self.wb = test_pivot_table.app1.books.open(test_sheet)
		self.pts = self.wb.sheets[0].pivot_tables
		src_range = self.wb.sheets[0].range('A1').expand()
		dest_range = self.wb.sheets[0].range('E5').expand()

		self.assertEqual(0, len(self.pts))
		self.pt = self.pts.add(src_range, xw.Range('E5'),
		row_fields = ['Name'], column_fields=['Age'],
		page_fields=['Arg1'], data_fields=['Arg2'])
		self.assertEqual(1, len(self.pts))

	def tearDown(self):
		for pt in self.pts:
			pt.delete()

	@classmethod
	def tearDownClass(cls):
		test_pivot_table.app1.quit()

				# def test_pivot_table_dest(self):
				# 	pass
                #

				# def test_pivot_table_src(self):
				# 	pass

	def test_name(self):
		prev_name = self.pt.name
		self.assertIsNot('New Name', prev_name)
		self.pt.name = 'New Name'
		self.assertEqual('New Name', self.pt.name)

	def test_correct_row_fields(self):
		self.assertEqual(['Name'], self.pt.row_fields)

	def test_correct_column_fields(self):
		self.assertEqual(['Age'], self.pt.column_fields)

	def test_correct_page_fields(self):
		self.assertEqual(['Arg1'], self.pt.page_fields)

	def test_correct_data_fields(self):
		self.assertEqual(['Sum of Arg2'], self.pt.data_fields)

	def test_edit_row_field(self):
		self.pt.hide_field('Name')
		self.assertEqual([], self.pt.row_fields)
		self.pt.add_row_field('Name')
		self.assertEqual(['Name'], self.pt.row_fields)

	def test_edit_column_field(self):
		self.pt.hide_field('Age')
		self.assertEqual([], self.pt.column_fields)
		self.pt.add_column_field('Age')
		self.assertEqual(['Age'], self.pt.column_fields)

	def test_edit_page_field(self):
		self.pt.hide_field('Arg1')
		self.assertEqual([], self.pt.page_fields)
		self.pt.add_page_field('Arg1')
		self.assertEqual(['Arg1'], self.pt.page_fields)

	def test_edit_data_field(self):
		self.pt.hide_field('Sum of Arg2')
		self.assertEqual([], self.pt.data_fields)
		self.pt.add_data_field('Arg2', 'SUM')
		self.assertEqual(['Sum of Arg2'], self.pt.data_fields)

	def test_delete(self):
		self.pt.delete()
		self.assertEqual(0, len(self.pts))


if __name__ == '__main__':
	unittest.main()
