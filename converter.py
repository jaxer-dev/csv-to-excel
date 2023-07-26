import os
import pandas as pd

def convert_csv_to_xlsx(csv_file):
    try:
        df = pd.read_csv(csv_file)

        max_rows_per_sheet = 1048575  # Maksymalna liczba wierszy na arkusz Excela

        num_sheets = (len(df) - 1) // max_rows_per_sheet + 1

        base_name = os.path.splitext(csv_file)[0]
        xlsx_file = f"{base_name}.xlsx"

        with pd.ExcelWriter(xlsx_file, engine='openpyxl') as writer:
            for i in range(num_sheets):
                start_idx = i * max_rows_per_sheet
                end_idx = min((i + 1) * max_rows_per_sheet, len(df))
                sheet_df = df.iloc[start_idx:end_idx]

                sheet_name = f"Sheet_{i+1}" if i > 0 else "Sheet"
                sheet_df.to_excel(writer, sheet_name=sheet_name, index=False)

        return True
    except Exception as e:
        print(f"Error converting {csv_file} to XLSX: {e}")
        return False

if __name__ == '__main__':
    folder_path = './'  # Zmień na ścieżkę do swojego folderu z plikami CSV

    csv_files = [os.path.join(folder_path, f) for f in os.listdir(folder_path) if f.endswith('.csv')]
    if not csv_files:
        print("Brak plików CSV w folderze.")
    else:
        print(f"Znaleziono {len(csv_files)} plików CSV. Rozpoczęcie konwersji na XLSX.")

        for csv_file in csv_files:
            if convert_csv_to_xlsx(csv_file):
                print(f"Pomyślnie przekonwertowano {csv_file} do XLSX.")
            else:
                print(f"Konwersja {csv_file} nie powiodła się.")

        print("Zakończono konwersję wszystkich plików.")
