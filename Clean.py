import re
import tkinter as tk
from tkinter import filedialog
from thefuzz import fuzz
import pandas as pd

# select excel file because there is nothing cool about pasting a directory
root = tk.Tk()
file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])

sheet_name = 'Cars'  # Name which sheet to read
df = pd.read_excel(file_path, sheet_name=sheet_name)

# Makes sure the column exists in the df
if 'Vehicle_Make' in df.columns:
    # 'Vehicle_Make' column to uppercase
    df['Vehicle_Make'] = df['Vehicle_Make'].str.replace(r'[^A-Za-z0-9\s]', '', regex=True).str.upper()
else:
    raise KeyError("The column specified does not exist. Check the name and try again.")

# Define Known Vehicle Make Dictionary.
known_make = {'TOYOTA', 'HONDA', 'CHEVROLET', 'NISSAN', 'FORD', 'DODGE', 'HYUNDAI', 'BMW',
              'MERCEDES-BENZ', 'VOLKSWAGEN', 'LEXUS', 'MAZDA', 'SUBARU', 'MITSUBISHI', 'JEEP', 'KIA', 'TESLA',
              'INFINITI', 'ACURA', 'CADILLAC', 'CHRYSLER', 'BUICK', 'AUDI', 'PORSCHE', 'LAND ROVER', 'RANGE ROVER'
              'MINI COOPER', 'VOLVO', 'FIAT',
              'ALFA ROMEO', 'JAGUAR', 'MASERATI', 'BENTLEY', 'ROLLS-ROYCE',
              'FERRARI', 'LAMBORGHINI', 'ASTON MARTIN', 'LOTUS', 'MCLAREN', 'BUGATTI', 'LUCID', 'GMC', 'RIVIAN'}


# Match "Vehicle_Make" to the known Makes
def make_match(Vehicle_Make, threshold=70):
    if Vehicle_Make in known_make:
        return Vehicle_Make, 100
    else:
        for make in known_make:
            score = fuzz.ratio(Vehicle_Make, make)
            if score > threshold:
                return make, score
        return Vehicle_Make, 0


# Apply make_match to the 'Vehicle_Make' col and create 'Matched_Make' and 'Score'
df[['Matched_Make', 'Score']] = df['Vehicle_Make'].apply(lambda x: pd.Series(make_match(x, threshold=70)))

# Select 20 rows and display results. Use random_state=1 for the same data to be used for reproducibility purposes.
random_selection = df.sample(n=20, random_state=1)
print(random_selection[['Matched_Make', 'Vehicle_Make', 'Score']])

# Missing or 'Unknown' values in the 'Vehicle_Make' column
df['Vehicle_Make'] = df['Vehicle_Make'].fillna('UNKNOWN').replace('Unknown', 'UNKNOWN')

# Select where to save file. Because you get to choose where to save your file.
file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
if file_path:
    df.to_excel(file_path, index=False)
    print(f"file saved at: {file_path}")

#%%
