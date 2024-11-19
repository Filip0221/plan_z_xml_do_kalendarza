import re

kolumna_dzien = {
    'C': 'Poniedziałek',
    'D': 'Poniedziałek',
    'E': 'Poniedziałek',
    'F': 'Poniedziałek',
    'G': 'Poniedziałek',
    'H': 'Poniedziałek',
    'I': 'Poniedziałek',
    'J': 'Wtorek',
    'K': 'Wtorek',
    'L': 'Wtorek',
    'M': 'Wtorek',
    'N': 'Wtorek',
    'O': 'Wtorek',
    'P': 'Wtorek',
    'Q': 'Środa',
    'R': 'Środa',
    'S': 'Środa',
    'T': 'Środa',
    'U': 'Środa',
    'V': 'Środa',
    'W': 'Środa',
    'X': 'Czwartek',
    'Y': 'Czwartek',
    'Z': 'Czwartek',
    'AA': 'Czwartek',
    'AB': 'Czwartek',
    'AC': 'Czwartek',
    'AD': 'Piątek',
    'AE': 'Piątek',
    'AF': 'Piątek',
    'AG': 'Piątek',
    'AH': 'Piątek',
    'AI': 'Piątek',
}

def get_dzien_from_kolumna(kolumna_komorki):
    return kolumna_dzien.get(kolumna_komorki, 'Nieznany dzień')

def get_column_letter(komorka):
    coordinate = komorka.coordinate
    if len(coordinate) >= 2 and coordinate[1].isalpha():  # Jeśli kolumna ma dwie litery, np. "AA10"
        return coordinate[:2]  # Zwracamy pierwsze dwie litery
    return coordinate[0]  # Zwracamy pierwszą literę dla kolumn jednoliterowych

def remove_between_zp_gr13(text):
    # Replace everything between (ZP) or (ĆW) and gr. 13 with an empty string
    return re.sub(r'(\(ZP\)|\(ćw\))\s*.*?gr\.\s*13', r'\1 gr. 13', text, flags=re.IGNORECASE)


def split_and_remove_after_gr13(text):
     # Dopasowanie fragmentu tekstu przed i włącznie z gr 13
    text = re.sub(r'13,', '13 ,', text)
    text = re.sub(r',13', ', 13', text)
    match = re.search(r'(.*?13)(.*)', text)
    if match:
        # Pierwsza część: przed gr 13, w tym gr 13 i ewentualnie inne grupy (np. gr 13,14)
        before_gr13 = match.group(1).strip()
        after_gr13 = match.group(2).strip()
        match = re.search(r'\((.*?)\)', after_gr13)
        if match:
            return before_gr13, match.group(1) 
        return before_gr13, after_gr13
    
    return text, text
