from datetime import datetime



class Termin:
    def __init__(self, data_zajec, tydzien, dzien_tygodnia):
        # Jeżeli data_zajec to tekst 'x' to przypisujemy None
        if data_zajec == 'x':
            self.data_zajec = None
        elif isinstance(data_zajec, str):
            # Jeżeli data jest tekstem, konwertujemy ją na datetime
            self.data_zajec = datetime.strptime(data_zajec, "%Y-%m-%d")
        else:
            self.data_zajec = data_zajec  # Jeżeli jest to już obiekt datetime
        self.tydzien = tydzien
        self.dzien_tygodnia = dzien_tygodnia

    def __str__(self):
        # Jeżeli data_zajec to None, to nie wyświetlamy daty
        if self.data_zajec:
            return f"Zajęcia: {self.data_zajec.strftime('%Y-%m-%d')}, Tydzień: {self.tydzien}, Dzień tygodnia: {self.dzien_tygodnia}"
        else:
            return f"Zajęcia: Brak, Tydzień: {self.tydzien}, Dzień tygodnia: {self.dzien_tygodnia}"


class Zajecia:
    def __init__(self, dzien, adres, wartosc, tygodnie, czas_rozpoczecia=None, czas_zakonczenia=None):
        self.dzien = dzien
        self.adres = adres
        self.wartosc = wartosc
        self.tygodnie = self.extract_weeks(tygodnie)
        self.czas_rozpoczecia = czas_rozpoczecia
        self.czas_zakonczenia = czas_zakonczenia

    def __str__(self):
        return f"Dzień: {self.dzien}, Adres: {self.adres}, Wartość: {self.wartosc},\nTygodnie: {self.tygodnie}, Czas rozpoczęcia: {self.czas_rozpoczecia}, Czas zakończenia: {self.czas_zakonczenia}"

    def extract_weeks(self, week_range):
        """Funkcja, która przetwarza zakres tygodni"""
        weeks = []
        week_range = week_range.replace(" tydz.", "").replace("tydz.", "")
        parts = week_range.split(',')
        
        for part in parts:
            if '-' in part:
                start, end = part.split('-')
                weeks.extend(range(int(start), int(end) + 1))
            else:
                weeks.append(int(part))
        
        weeks = sorted(set(weeks))
        return weeks
