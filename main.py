import pandas as pd
import os

input_file = 'test_promo.xlsx'
output_file = 'result_promo_2.xlsx'

def main():
    if not os.path.exists(input_file):
        print(f"File {input_file} tidak ditemukan di {os.getcwd()}")
        return

    try:
        # Load data
        df = pd.read_excel(input_file)
        
        if 'PLU' not in df.columns:
            print("Header 'PLU' tidak ditemukan.")
            return

        # Data Cleaning & Transformation
        df['PLU'] = df['PLU'].astype(str).str.replace(r'\.0$', '', regex=True)
        
        # Unpivot Matrix ke Long Format
        melted = df.melt(id_vars=['PLU'], var_name='Store', value_name='Status')
        
        # Filter & Export
        result = melted[melted['Status'].astype(str).str.upper() == 'AKTIF'].copy()
        result[['PLU', 'Store']].to_excel(output_file, index=False)
        
        print(f"Berhasil: {len(result)} baris diproses.")
        print(result[['PLU', 'Store']].head())

    except Exception as e:
        print(f"Error: {e}")

if __name__ == "__main__":
    main()