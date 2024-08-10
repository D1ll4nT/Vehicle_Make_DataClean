import unittest
from unittest.mock import patch
import pandas as pd
from thefuzz import fuzz

class TestVehicleMakeMatching(unittest.TestCase):

    @patch('tkinter.filedialog.askopenfilename')
    @patch('tkinter.filedialog.asksaveasfilename')
    def test_known_make_match(self, mock_saveas, mock_open):
        mock_open.return_value = 'test.xlsx'
        mock_saveas.return_value = 'output.xlsx'
        df = pd.DataFrame({'Vehicle_Make': ['TOYOTA', 'HONDA', 'GMC']})
        df['Vehicle_Make'] = df['Vehicle_Make'].str.replace(r'[^A-Za-z0-9\s]', '', regex=True).str.upper()
        known_make = {'TOYOTA', 'HONDA', 'GMC'}
        def make_match(Vehicle_Make, threshold=70):
            if Vehicle_Make in known_make:
                return Vehicle_Make, 100
            else:
                for make in known_make:
                    score = fuzz.ratio(Vehicle_Make, make)
                    if score > threshold:
                        return make, score
                return Vehicle_Make, 0
        df[['Matched_Make', 'Score']] = df['Vehicle_Make'].apply(lambda x: pd.Series(make_match(x, threshold=70)))
        self.assertEqual(df.loc[0, 'Score'], 100)
        self.assertEqual(df.loc[1, 'Score'], 100)
        self.assertEqual(df.loc[2, 'Score'], 100)

    @patch('tkinter.filedialog.askopenfilename')
    @patch('tkinter.filedialog.asksaveasfilename')
    def test_unknown_make_match(self, mock_saveas, mock_open):
        mock_open.return_value = 'test.xlsx'
        mock_saveas.return_value = 'output.xlsx'
        df = pd.DataFrame({'Vehicle_Make': ['UNKNOWNMAKE']})
        df['Vehicle_Make'] = df['Vehicle_Make'].str.replace(r'[^A-Za-z0-9\s]', '', regex=True).str.upper()
        known_make = {'TOYOTA', 'HONDA', 'GMC'}
        def make_match(Vehicle_Make, threshold=70):
            if Vehicle_Make in known_make:
                return Vehicle_Make, 100
            else:
                for make in known_make:
                    score = fuzz.ratio(Vehicle_Make, make)
                    if score > threshold:
                        return make, score
                return Vehicle_Make, 0
        df[['Matched_Make', 'Score']] = df['Vehicle_Make'].apply(lambda x: pd.Series(make_match(x, threshold=70)))
        self.assertEqual(df.loc[0, 'Score'], 0)

    @patch('tkinter.filedialog.askopenfilename')
    @patch('tkinter.filedialog.asksaveasfilename')
    def test_special_characters_removal(self, mock_saveas, mock_open):
        mock_open.return_value = 'test.xlsx'
        mock_saveas.return_value = 'output.xlsx'
        df = pd.DataFrame({'Vehicle_Make': ['T@oyota!', 'H#onda$', 'G%MC^']})
        df['Vehicle_Make'] = df['Vehicle_Make'].str.replace(r'[^A-Za-z0-9\s]', '', regex=True).str.upper()
        self.assertEqual(df.loc[0, 'Vehicle_Make'], 'TOYOTA')
        self.assertEqual(df.loc[1, 'Vehicle_Make'], 'HONDA')
        self.assertEqual(df.loc[2, 'Vehicle_Make'], 'GMC')

if __name__ == '__main__':
    unittest.main()