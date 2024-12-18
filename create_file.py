import pandas as pd

# Create a sample DataFrame
data = {
    'Name': ['Project Reignite', 'Jinshan Songlin Market Swine Manure Utilization Project'],
    'Location': ['Tanzania', 'China'],
    'Developer': ['EcoProjects Ltd', 'GreenEnergy Co.']
}

df = pd.DataFrame(data)

# Save the DataFrame to an Excel file
df.to_excel('your_file.xlsx', index=False, sheet_name='Results')

print("Sample Excel file created successfully.")
