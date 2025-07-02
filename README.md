# Zaawansowany filtr danych w Excelu 

## Opis projektu

Projekt przedstawia narzędzie do zaawansowanego filtrowania danych w programie Microsoft Excel, stworzone z myślą o analitykach i osobach pracujących na dużych zbiorach danych. Narzędzie zostało oparta o formularze, przyciski oraz kod VBA, umożliwiając użytkownikowi filtrowanie danych po wielu warunkach jednocześnie – bez konieczności znajomości formuł, Power Query czy kodowania.

## Funkcjonalności

- **Automatyczne przygotowanie danych**
  - Użytkownik wkleja tabelę do wskazanego arkusza.
  - Po kliknięciu przycisku „PRZYGOTUJ” dane zostają przeniesione do głównego arkusza filtrowania.

- **Filtrowanie po wielu kolumnach**
  - Możliwość ustawienia niezależnych warunków dla wielu kolumn równocześnie.
  - Obsługa filtrów logicznych i warunków porównawczych (np. `>1000`, `<2024-01-01`, `=Warszawa`, `<>0`, `>=50 <100`).

- **Resetowanie widoku**
  - Przycisk „RESET” umożliwia przywrócenie danych do pełnego, niefiltrowanego stanu.
  - Dostępny również przycisk czyszczący nagłówki filtra.

- **Czytelny interfejs i instrukcja**
  - Wbudowana instrukcja obsługi oraz arkusz pomocy z przykładami warunków filtracji.
  - Intuicyjne rozmieszczenie przycisków i pól filtrowania.

## Struktura pliku

- `Instrukcja` – szczegółowy opis działania narzędzia krok po kroku
- `Przygotowanie` – arkusz wejściowy do wklejenia tabeli źródłowej
- `Dane` – główny arkusz roboczy do filtrowania danych
- `Ściąga warunków do filtra` – przykłady dopuszczalnych warunków filtracyjnych

## Jak korzystać

1. Wklej dane tabelaryczne do arkusza „Przygotowanie”.
2. Kliknij przycisk „PRZYGOTUJ”, aby przygotować dane.
3. W arkuszu „Dane” ustaw warunki filtrowania pod odpowiednimi nagłówkami.
4. Kliknij przycisk „FILTRUJ”, aby zastosować wybrane filtry.
5. Użyj przycisku „RESET”, aby przywrócić pełny widok danych.
6. W każdej chwili możesz wyczyścić nagłówki filtracyjne.

## Technologia

- Microsoft Excel (plik .xlsm)
- Visual Basic for Applications (VBA)
- Formuły Excelowe (warunki logiczne, porównania)

## Zastosowania

- Filtrowanie danych sprzedażowych, kadrowych, finansowych
- Przygotowanie danych do analiz i raportów
- Usprawnienie codziennej pracy z dużymi tabelami
- Narzędzie wewnętrzne do procesów raportowych

## Potencjalny rozwój

- Zapis i odczyt zestawów filtrów
- Eksport wyników do osobnego arkusza lub pliku
- Możliwość filtrowania z wykorzystaniem warunków logicznych „OR”
- Walidacja wpisywanych warunków i komunikaty błędów

Ten plik Excel został opracowany przeze mnie jako uniwersalne, elastyczne narzędzie do zaawansowanego filtrowania danych, szczególnie przydatne dla analityków, specjalistów ds. raportowania oraz każdego, kto regularnie pracuje z dużymi zestawami danych w Excelu. Projekt umożliwia filtrowanie danych po wielu kolumnach jednocześnie, z zastosowaniem złożonych warunków logicznych, dzięki czemu pozwala znacząco skrócić czas analizy i zwiększyć precyzję pracy.
Zaprojektowany interfejs jest intuicyjny – użytkownik wkleja swoją tabelę, definiuje warunki filtrów i jednym kliknięciem otrzymuje przefiltrowany wynik. Całość została zautomatyzowana z wykorzystaniem języka VBA, dzięki czemu użytkownik nie musi znać kodu, aby korzystać z funkcjonalności.
Ten projekt to część mojego portfolio, w którym prezentuję narzędzia ułatwiające codzienną pracę z danymi. Jeśli chcesz zaimplementować podobne rozwiązanie w swojej organizacji lub potrzebujesz wersji dostosowanej do konkretnego procesu.

![image](https://github.com/user-attachments/assets/9f7de2b4-fbcd-464a-8906-e3f222aebb39)
![image](https://github.com/user-attachments/assets/c08e07a0-19fd-451e-9a19-818bfb2cdceb)
![image](https://github.com/user-attachments/assets/54835eea-19eb-497a-95b4-3d5297464061)



