import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment

# Original data
data = {
    'Name' : ['a','b','c','d','e','f','g','h','i','j','k','l','m','n','o','p','q','r','s','t'],
    'Roll' : list(range(1, 21)),
    'id' : list(range(101, 121)),
    'city' : ['A','B','C','D','E','G','H','I','J','K','L','M','N','O','P','Q','R','S','T','U'],
    'class' : [f'cse-{i}' for i in range(1, 21)]
}

df = pd.DataFrame(data)

# Save dataframe to Excel
excel_path = "student.xlsx"
df.to_excel(excel_path, index=False)

print("Excel file 'student.xlsx' created successfully!")
