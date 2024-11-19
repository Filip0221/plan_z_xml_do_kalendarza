import openpyxl
from zajecia import Zajecia
from funkcje import get_column_letter, remove_between_zp_gr13, split_and_remove_after_gr13, get_dzien_from_kolumna
import re


# Ścieżka do pliku Excel
plik_excel = 'C:/Users/filip.skup/Desktop/harmonogram.xlsx'
#plik_excel = '/Users/filipskup/Desktop/plan_z_xml_do_kalendarza/harmonogram.xlsx'
# Otwórz plik Excel
wb = openpyxl.load_workbook(plik_excel)

# Wybierz arkusz
arkusz = wb['fiz IV']

zajecia_list = []  # lista do przechowywania obiektów zajęć



for kolumna in arkusz.iter_cols(min_row=8, max_row=67, min_col=3, max_col=36):
    for komorka in kolumna:
        if komorka.value is not None:
            komorka_value = komorka.value.strip().lower()
            if re.search(r'\bgr\.\s*11\s*,\s*13\b|\bgr\.\s*13\b|\bgr\s*13\b', komorka_value):
                kolumna_komorki = get_column_letter(komorka)
                dzien = get_dzien_from_kolumna(kolumna_komorki)
                


                komorka_value = remove_between_zp_gr13(komorka_value)
                # komorka_value bez niczego po gr 13
                komorka_value, tygodnie = split_and_remove_after_gr13(komorka_value)
                
                # Tworzenie obiektu Zajęcia i dodawanie do listy

                zajecie = Zajecia(dzien, komorka.coordinate, komorka_value, tygodnie)
                zajecia_list.append(zajecie)

                
for zajecie in zajecia_list:
    print(zajecie)
