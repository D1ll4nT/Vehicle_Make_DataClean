
# CleanVehicle_Make Script

This document provides a step-by-step explanation of the `CleanVehicle_Make` script, which is used for cleaning and matching vehicle make names in an Excel sheet.

## Overview

The `CleanVehicle_Make` script processes an Excel file containing vehicle make data, standardizes the make names, and attempts to match them against a predefined list of known vehicle makes using fuzzy matching. The results are then saved to a new Excel file.

## Instructions

### 1. Import Libraries

The script uses the following libraries:
- `re`: For regular expression operations.
- `tkinter`: For creating file dialogs to select and save files.
- `thefuzz`: For fuzzy string matching.
- `pandas`: For data manipulation and analysis.

### 2. Select Excel File

The script opens a file dialog to allow the user to select an Excel file:

```python
root = tk.Tk()
file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
```

### 3. Load Data

The script reads the specified sheet (`'Cars'`) from the selected Excel file into a pandas DataFrame:

```python
sheet_name = 'Cars'  # Name which sheet to read
df = pd.read_excel(file_path, sheet_name=sheet_name)
```

### 4. Clean Vehicle Make Column

The script checks if the `Vehicle_Make` column exists and cleans the data by removing non-alphanumeric characters and converting the text to uppercase:

```python
if 'Vehicle_Make' in df.columns:
    df['Vehicle_Make'] = df['Vehicle_Make'].str.replace(r'[^A-Za-z0-9\s]', '', regex=True).str.upper()
else:
    raise KeyError("The column specified does not exist. Check the name and try again.")
```

### 5. Define Known Vehicle Make Dictionary

The script contains a predefined dictionary of known vehicle makes:

```python
known_make = {'TOYOTA', 'HONDA', 'CHEVROLET', 'NISSAN', 'FORD', 'DODGE', 'HYUNDAI', 'BMW',
              'MERCEDES-BENZ', 'VOLKSWAGEN', 'LEXUS', 'MAZDA', 'SUBARU', 'MITSUBISHI', 'JEEP', 'KIA', 'TESLA',
              'INFINITI', 'ACURA', 'CADILLAC', 'CHRYSLER', 'BUICK', 'AUDI', 'PORSCHE', 'LAND ROVER', 'RANGE ROVER'
              'MINI COOPER', 'VOLVO', 'FIAT',
              'ALFA ROMEO', 'JAGUAR', 'MASERATI', 'BENTLEY', 'ROLLS-ROYCE',
              'FERRARI', 'LAMBORGHINI', 'ASTON MARTIN', 'LOTUS', 'MCLAREN', 'BUGATTI', 'LUCID', 'GMC', 'RIVIAN'}
```

### 6. Match Vehicle Make

The script defines a function `make_match` that attempts to match the `Vehicle_Make` against the known makes using a fuzzy matching algorithm:

```python
def make_match(Vehicle_Make, threshold=70):
    if Vehicle_Make in known_make:
        return Vehicle_Make, 100
    else:
        for make in known_make:
            score = fuzz.ratio(Vehicle_Make, make)
            if score > threshold:
                return make, score
        return Vehicle_Make, 0
```

### 7. Apply Matching

The script applies the `make_match` function to the `Vehicle_Make` column and creates new columns `Matched_Make` and `Score`:

```python
df[['Matched_Make', 'Score']] = df['Vehicle_Make'].apply(lambda x: pd.Series(make_match(x, threshold=70)))
```

### 8. Display Sample Results

A random selection of 20 rows is displayed to show the results of the matching process:

```python
random_selection = df.sample(n=20, random_state=1)
print(random_selection[['Matched_Make', 'Vehicle_Make', 'Score']])
```

### 9. Handle Missing Values

The script replaces any missing or unknown values in the `Vehicle_Make` column with `'UNKNOWN'`:

```python
df['Vehicle_Make'] = df['Vehicle_Make'].fillna('UNKNOWN').replace('Unknown', 'UNKNOWN')
```

### 10. Save Processed File

The script prompts the user to select a location to save the processed file:

```python
file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
if file_path:
    df.to_excel(file_path, index=False)
    print(f"file saved at: {file_path}")
```

## Feedback and Support

For any issues or suggestions, please reach out through the appropriate channels.

---

You can use the above documentation to understand and work with the `CleanVehicle_Make` script. If you need further assistance or want to modify the script, feel free to ask.
