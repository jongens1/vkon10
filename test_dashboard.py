import unittest
from pathlib import Path
import openpyxl
from streamlit.testing.v1 import AppTest
from app import normalize_pivot, aggregate, METRICS


class DashboardTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        workbook = openpyxl.load_workbook(Path(__file__).with_name('datax.xlsx'), read_only=True, data_only=True)
        cls.rows = list(workbook['DATA'].iter_rows(values_only=True))
        workbook.close()

    def test_actual_workbook_and_reconciliation(self):
        created, picked, unknown, info = normalize_pivot(self.rows)
        self.assertEqual(info['Načítané riadky'], 192)
        self.assertFalse(info['Neplatné hodnoty / nesúlad súčtov'])
        self.assertFalse(info['Nenamapované stĺpce'])
        self.assertFalse(info['Preskočené riadky'])
        self.assertEqual(created[METRICS].sum().tolist(), [1067702, 2294972])
        for m in METRICS:
            self.assertEqual(created[m].sum(), picked[m].sum() + unknown[m].sum())
        self.assertGreater(picked['Čas'].min().year, 1900)
        hourly = aggregate(created, picked, created['Čas'].min().date(), picked['Čas'].max().date(), 'h')
        daily = aggregate(created, picked, created['Čas'].min().date(), picked['Čas'].max().date(), 'D')
        self.assertEqual(hourly.iloc[:, 1:].sum().tolist(), daily.iloc[:, 1:].sum().tolist())

    def test_shifted_headers_and_repeated_labels(self):
        shifted = [[None] * (len(self.rows[0]) + 2)] * 3 + [[None, None, *r] for r in self.rows]
        original = normalize_pivot(self.rows)
        changed = normalize_pivot(shifted)
        for a, b in zip(original[:3], changed[:3]):
            self.assertTrue(a.equals(b))

    def test_merged_style_headers_and_invalid_value(self):
        rows = [list(r) for r in self.rows]
        for r in (5, 6):
            for c in range(len(rows[r]) - 1, 2, -1):
                if rows[r][c] == rows[r][c - 1]:
                    rows[r][c] = None
        self.assertTrue(normalize_pivot(rows)[0].equals(normalize_pivot(self.rows)[0]))
        rows[8][2] = 'broken'
        self.assertIn('C9', normalize_pivot(rows)[3]['Neplatné hodnoty / nesúlad súčtov'])

    def test_app_and_filters(self):
        app = AppTest.from_file(str(Path(__file__).with_name('app.py'))).run(timeout=60)
        self.assertFalse(app.exception)
        self.assertFalse(app.error)
        self.assertEqual(len(app.metric), 6)
        self.assertEqual(len(app.get('plotly_chart')), 2)
        app.sidebar.selectbox[0].set_value('deň').run()
        self.assertFalse(app.exception)
        from datetime import date
        app.sidebar.date_input[0].set_value((date(2026, 10, 5), date(2026, 10, 5))).run()
        self.assertFalse(app.exception)
        self.assertFalse(app.error)
        self.assertEqual(len(app.dataframe[-1].value), 1)


if __name__ == '__main__':
    unittest.main()
