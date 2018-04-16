import sys
sys.path.append(r'/Users/Harrison/anaconda3/envs/XLWingsDev/xlwings/')
import xlwings as xw
from xlwings.tests.common import TestBase, this_dir
import unittest
from parameterized import parameterized

test_sheet = 'test_pivot.xlsx'
SPEC = '/Applications/Microsoft Office 2011/Microsoft Excel.app'


class test_pivot_table(unittest.TestCase):

	@classmethod
	def setUpClass(cls):
		cls.app1 = xw.App(visible=False, spec=SPEC)
		cls.wb = cls.app1.books.open(test_sheet)
		cls.sht = cls.wb.sheets[0]

	def setUp(self):
		self.pts = test_pivot_table.sht.pivot_tables
		src_range = test_pivot_table.sht.range('A1').expand()
		dest_range = test_pivot_table.sht.range('E5').expand()

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

	@parameterized.expand([('Sum of Arg2', 'SUM'), ('Count of Arg2', 'COUNT'),
					('Average of Arg2', 'AVERAGE'), ('StdDev of Arg2', 'STD_DEV'),
					('StdDevp of Arg2', 'STD_DEV_P'), ('Var of Arg2', 'VAR'),
					('Varp of Arg2', 'VAR_P'), ('Count of Arg2', 'COUNT_NUMS')])
	def test_edit_data_field(self, data_field_string, data_field_func):
		self.pt.hide_field('Sum of Arg2')
		self.assertEqual([], self.pt.data_fields)
		self.pt.add_data_field('Arg2', data_field_func)
		self.assertEqual([data_field_string], self.pt.data_fields)

		self.pt.hide_field(data_field_string)
		self.assertEqual([], self.pt.data_fields)

	def test_delete(self):
		self.pt.delete()
		self.assertEqual(0, len(self.pts))

	def test_hide_field(self):
		self.pt.hide_field('Age')
		assert 'Age' not in self.pt.column_fields

		self.pt.hide_field('Sum of Arg2')
		assert 'Sum of Arg2' not in self.pt.data_fields

	def set_row_fields(self):
		pass

	def set_column_fields(self):
		pass

	def set_page_fields(self):
		pass

	def set_data_fields(self):
		pass


if __name__ == '__main__':
	unittest.main()
