class Zajecia:
    def __init__(self, dzien, adres, wartosc, tygodnie, czas_rozpoczecia=None, czas_zakonczenia=None):
        self.dzien = dzien
        self.adres = adres
        self.wartosc = wartosc
        self.tygodnie = self.extract_weeks(tygodnie)
        self.czas_rozpoczecia = czas_rozpoczecia
        self.czas_zakonczenia = czas_zakonczenia

    def __str__(self):
        return f"Dzień: {self.dzien}, Adres: {self.adres}, Wartość: {self.wartosc},\n Tygodnie: {self.tygodnie}, Czas rozpoczęcia: {self.czas_rozpoczecia}, Czas zakończenia: {self.czas_zakonczenia}"

    def extract_weeks(self, week_range):
        """Funkcja, która przetwarza zakres tygodni"""
        weeks = []
        week_range = week_range.replace(" tydz.", "").replace("tydz.", "")
        # Rozdzielenie na różne części (np. "1-6", "10", "11")
        parts = week_range.split(',')
        
        for part in parts:
            # Sprawdzamy, czy część to zakres tygodni (np. 1-6)
            if '-' in part:
                start, end = part.split('-')
                # Dodajemy wszystkie tygodnie w tym zakresie do listy
                weeks.extend(range(int(start), int(end) + 1))
            else:
                # Jeśli część to pojedynczy tydzień (np. 10)
                weeks.append(int(part))
        
        # Usuwamy duplikaty i sortujemy tygodnie
        weeks = sorted(set(weeks))
        return weeks

