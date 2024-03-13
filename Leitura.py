import pandas as pd

# Substitua pelo caminho correto para a sua Planilha Clínica
clinica_path = r'C:\Users\snatanael\Downloads\Planilha Clinica.xlsx'
clinica_df = pd.read_excel(clinica_path)
print(clinica_df.columns.tolist())
