import unittest
from financial_engine import calculate_pv, generate_ifrs_schedule, generate_us_gaap_schedule, perform_comparison

class TestLeaseFinancials(unittest.TestCase):
    
    def setUp(self):
        self.pmt = 10000.0
        self.rate = 0.05
        self.term = 24
        
    def test_pv_calculation(self):
        # Math verification for PMT=$10,000, rate=5%, n=24
        computed_pv = calculate_pv(self.pmt, self.rate, self.term)
        self.assertAlmostEqual(computed_pv, 227941.02, places=2)
        
    def test_ifrs_amortization(self):
        schedule = generate_ifrs_schedule(self.pmt, self.rate, self.term)
        
        # Month 0 check
        self.assertEqual(schedule[0]['month'], 0)
        self.assertEqual(schedule[0]['liability_closing'], 227941.02)
        self.assertEqual(schedule[0]['rou_closing'], 227941.02)
        
        # Month 1 check
        # interest = round(227941.02 * 0.05 / 12, 2) = 949.75
        # depreciation = round(227941.02 / 24, 2) = 9497.54
        # closing_liab = 227941.02 - 10000 + 949.75 = 218890.77
        # closing_rou = 227941.02 - 9497.54 = 218443.48
        m1 = schedule[1]
        self.assertEqual(m1['month'], 1)
        self.assertEqual(m1['interest'], 949.75)
        self.assertEqual(m1['depreciation'], 9497.54)
        self.assertEqual(m1['liability_closing'], 218890.77)
        self.assertEqual(m1['rou_closing'], 218443.48)
        
        # Month 12 check (Year 1 end)
        m12 = schedule[12]
        self.assertEqual(m12['liability_closing'], 116814.35)
        self.assertEqual(m12['rou_closing'], 113970.54)
        
        # Month 24 check (Lease end - should close to 0)
        m24 = schedule[24]
        self.assertEqual(m24['liability_closing'], 0.0)
        self.assertEqual(m24['rou_closing'], 0.0)
        
    def test_us_gaap_amortization(self):
        schedule = generate_us_gaap_schedule(self.pmt, self.rate, self.term)
        
        # Month 0 check
        self.assertEqual(schedule[0]['month'], 0)
        self.assertEqual(schedule[0]['liability_balance'], 227941.02)
        self.assertEqual(schedule[0]['rou_balance'], 227941.02)
        
        # Month 1 check
        # lease expense = 10000
        # interest = 949.75
        # liab_reduction = 10000 - 949.75 = 9050.25
        # rou_amortization = 9050.25
        # liability_balance = 227941.02 - 9050.25 = 218890.77
        # rou_balance = 227941.02 - 9050.25 = 218890.77
        m1 = schedule[1]
        self.assertEqual(m1['month'], 1)
        self.assertEqual(m1['interest'], 949.75)
        self.assertEqual(m1['liab_reduction'], 9050.25)
        self.assertEqual(m1['rou_amortization'], 9050.25)
        self.assertEqual(m1['liability_balance'], 218890.77)
        self.assertEqual(m1['rou_balance'], 218890.77)
        
        # Month 12 check (Year 1 end)
        m12 = schedule[12]
        self.assertEqual(m12['liability_balance'], 116814.35)
        self.assertEqual(m12['rou_balance'], 116814.35)
        
        # Month 24 check (Lease end - should close to 0)
        m24 = schedule[24]
        self.assertEqual(m24['liability_balance'], 0.0)
        self.assertEqual(m24['rou_balance'], 0.0)
        
    def test_perform_comparison(self):
        results = perform_comparison(self.pmt, self.rate, self.term)
        
        # Year 1 (Month 12 closing) totals comparison
        y1 = results['yearly_comparison'][0]
        self.assertEqual(y1['year'], 1)
        
        # IFRS expense Year 1 = Sum of interest (8873.33) + Sum of dep (113970.48) = 122843.81
        self.assertEqual(y1['ifrs_expense'], 122843.81)
        # US GAAP expense Year 1 = 120000
        self.assertEqual(y1['us_expense'], 120000)
        # Expense diff = 122843.81 - 120000 = 2843.81
        self.assertEqual(y1['expense_diff'], 2843.81)
        
        # ROU Asset difference = 113970.54 (IFRS) - 116814.35 (US GAAP) = -2843.81
        self.assertEqual(y1['rou_diff'], -2843.81)
        # Lease Liability difference = 0.00
        self.assertEqual(y1['liab_diff'], 0.00)
        
        # Journal Entries count
        # Should have Month 0 (initial recognition), Year 1, Year 2
        self.assertEqual(len(results['journal_entries']), 3)
        
        # Check initial recognition entry
        init_je = results['journal_entries'][0]
        self.assertEqual(init_je['entries'][0]['account'], 'Right-of-Use (ROU) Asset')
        self.assertEqual(init_je['entries'][0]['debit'], 227941.02)
        
        # Check year 1 adjustment entry
        y1_je = results['journal_entries'][1]
        self.assertTrue(any('Retained Earnings' in e['account'] or 'Expense' in e['account'] for e in y1_je['entries']))
        self.assertTrue(any('Right-of-Use (ROU) Asset' in e['account'] for e in y1_je['entries']))

if __name__ == '__main__':
    unittest.main()
