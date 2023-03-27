import pandas as pd

# Read the text file using pandas
df = pd.read_csv('Arya7.txt', delimiter='\t', header=None)

# Save the data to an Excel file
df.to_excel('attendence.xlsx', index=False, header=False)
