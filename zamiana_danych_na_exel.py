import openpyxl
from openpyxl.styles import Alignment
from datetime import datetime

# Funkcja do parsowania danych
def parse_input_data(input_data):
    tygodnie = []
    lines = input_data.split("\n")
    for line in lines:
        line = line.strip()
        if line:
            # Używamy split() bez argumentu, co dzieli po każdej spacji
            parts = line.split()  
            if len(parts) > 1:
                tygodnie.append(parts)
    return tygodnie

# Funkcja do konwersji tekstu na obiekt datetime
def convert_to_date(date_str):
    try:
        return datetime.strptime(date_str, '%d.%m.%Y')  # Przekształcamy tekst w format daty
    except ValueError:
        return None  # Jeśli nie uda się przekonwertować, zwrócimy None

# Funkcja do konwersji danych na tabelę
def create_excel_table_from_input(data, filename="Harmonogram.xlsx"):
    wb = openpyxl.Workbook()  # Tworzenie nowego workbooka
    arkusz = wb.active
    arkusz.title = "Semestr zimowy 2024_2025"  # Zmieniona nazwa arkusza
    
    # Nagłówki tabeli
    naglowki = ["Tygodnie", "Poniedziałek", "Wtorek", "Środa", "Czwartek", "Piątek", "Sobota", "Niedziela"]
    arkusz.append(naglowki)  # Dodajemy nagłówki
    
    # Przetworz dane i dodaj je do arkusza
    for tydzien in data:
        tygodnie = tydzien[0]  # Tygodnie (np. '1')
        dni = tydzien[1:]  # Listę dni (np. daty)
        # Konwertujemy daty z tekstu na obiekty datetime
        dni_z_datami = [convert_to_date(d) if convert_to_date(d) else d for d in dni]
        rzad = [tygodnie] + dni_z_datami  # Dodajemy tygodnie jako pierwszy element
        arkusz.append(rzad)  # Dodajemy dane do arkusza
    
    # Formatowanie komórek
    for row in arkusz.iter_rows(min_row=1, max_row=len(data) + 1, min_col=2, max_col=8):
        for cell in row:
            if isinstance(cell.value, datetime):  # Jeżeli komórka zawiera datę
                cell.number_format = 'DD.MM.YYYY'  # Ustawiamy format daty
            cell.alignment = Alignment(horizontal='center', vertical='center')  # Centruj zawartość komórek

    # Zapisz plik
    wb.save(filename)
    print(f"Plik Excel został zapisany jako {filename}")

    
def create_exel ():
    # Wklej dane z PDF do zmiennej input_data
    input_data = """
    1 07.10.2024 01.10.2024 02.10.2024 03.10.2024 04.10.2024 05.10.2024 06.10.2024
    2 14.10.2024 08.10.2024 09.10.2024 10.10.2024 11.10.2024 12.10.2024 13.10.2024
    3 21.10.2024 15.10.2024 16.10.2024 17.10.2024 18.10.2024 19.10.2024 20.10.2024
    4 28.10.2024 22.10.2024 23.10.2024 24.10.2024 25.10.2024 26.10.2024 27.10.2024
    5 04.11.2024 05.11.2024 06.11.2024 07.11.2024 29.10.2024 x x
    6 08.11.2024 12.11.2024 13.11.2024 14.11.2024 15.11.2024 16.11.2024 17.11.2024
    7 18.11.2024 19.11.2024 20.11.2024 21.11.2024 22.11.2024 23.11.2024 24.11.2024
    8 25.11.2024 26.11.2024 27.11.2024 28.11.2024 29.11.2024 30.11.2024 01.12.2024
    9 02.12.2024 03.12.2024 04.12.2024 05.12.2024 06.12.2024 07.12.2024 08.12.2024
    10 09.12.2024 10.12.2024 11.12.2024 12.12.2024 13.12.2024 14.12.2024 15.12.2024
    11 16.12.2024 17.12.2024 18.12.2024 19.12.2024 20.12.2024 21.12.2024 22.12.2024
    12 13.01.2025 07.01.2025 08.01.2025 09.01.2025 10.01.2025 11.01.2025 12.01.2025
    13 20.01.2025 14.01.2025 15.01.2025 16.01.2025 17.01.2025 18.01.2025 19.01.2025
    14 27.01.2025 21.01.2025 22.01.2025 23.01.2025 24.01.2025 25.01.2025 26.01.2025
    15 03.02.2025 28.01.2025 29.01.2025 30.01.2025 31.01.2025 01.02.2025 02.02.2025
    """

    # Parsowanie danych
    parsed_data = parse_input_data(input_data)

    # Tworzenie tabeli w Excelu
    create_excel_table_from_input(parsed_data)
