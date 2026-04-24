# Notatki ze spotkania — AI_next2

**Źródło:** `AI_next2.vtt` (Teams auto-transkrypcja, ~47 min)
**Uczestnicy:** Robert Godziszewski + Paweł (drugi głos w tle)
**Kontekst:** Kontynuacja rozmowy strategicznej z AI_next1 — planowanie oferty AI, produktu szkoleniowego i przygotowanie do konferencji w Lizbonie.

---

## 1. Struktura oferty AI — składowe produktu

Rozmowa zaczyna się od syntezy wniosków z wcześniejszych dyskusji i próby zdefiniowania, co dokładnie sprzedawać.

Wyróżnione poziomy oferty:

| Poziom | Zakres |
|---|---|
| **Strategiczny** | Pierwsze 12–18 miesięcy strategii AI — od czego zacząć, na czym się skupić |
| **Taktyczny** | Projektowanie procesów pod AI; identyfikacja gdzie AI ma sens, a gdzie nie |
| **Operacyjny** | Automatyzacja przez makra, Power Automate, M365 tools, budowanie asystentów / agentów |

Wniosek: zamiast sprzedawać to jako trzy osobne produkty, należy opakować to w jedną narrację, którą będzie można elastycznie dopasowywać do poziomu klienta (GPO z dużym władztwem procesowym vs. jednostka z małym wpływem decyzyjnym).

---

## 2. Freelancer — budowanie sieci zewnętrznych trenerów

Krystian zasugerował poszukanie freelancera do realizacji szkoleń i projektów.

**Problem z historycznym podejściem:** Poprzednio, gdy pojawiała się potrzeba, skupiali się na realizacji zamiast sprzedaży. Gdy freelancer był potrzebny — brakowało kogoś gotowego. Vicious cycle.

**Planowane podejście:**
- Zidentyfikować 2–3 osoby, z którymi dobrze się pracuje, i trzymać je "pod ręką"
- Nie musi być od razu kontrakt — wystarczy nieformalna umowa: "jak będziemy mieli coś, to się odezwiemy"
- Profil szukanej osoby: podobny do Szymona — ktoś elastyczny, bez sztywnych zobowiązań, chętny do szkoleń ad hoc i ewentualnie do projektów

**Akcja dla Pawła:** Sprawdzić w swoich grupach (Development, warszawskie sieci), czy jest ktoś młody zainteresowany freelancingiem szkoleniowym.

---

## 3. Sezon wakacyjny — strategia szkoleń

Robert nie zgadza się z przekonaniem Krystiana, że wakacje to martwy sezon:

- Lipiec–sierpień to "ogórek" — mniej ciśnienia operacyjnego w firmach
- To dobry moment na szkolenia, bo ludzie mają więcej przestrzeni
- Wyzwanie: część pracowników na urlopach, ale generalnie spokojniej
- Marketing i aktywność muszą trwać przez cały rok — **nie można zamrażać działań na lato**

**Plan:** Zorganizować 2 webinary w sierpniu, żeby "rozgrzać" rynek przed nową falą sprzedaży od września.

---

## 4. Akcje priorytetowe — "rozbieranie słonia"

Zidentyfikowane trzy najszybsze działania do realizacji:

### 4.1 Aktualizacja prezentacji sprzedażowej AI Trainings

- Przepisać i odświeżyć ofertę szkoleń
- Zmienić nazwę "Intermediate" → **"Practice Lab"**
- Wokół tego zbudować mailing z wynikami ankiet satysfakcji uczestników
- Przekaz marketingowy: "Wasi ludzie używają AI tylko do pisania maili — bo nie wiedzą, co innego można z nim zrobić. Przyjdźcie na warsztaty i zobaczcie."

### 4.2 Budowanie Practice Lab opartego o dane z ankiet (Slido)

- Przejrzeć wyniki ankiet szkoleniowych pod kątem tego, czego uczestnikom brakowało
- Wszystko, czego brakowało na Basic'u → powinno trafić do Practice Lab
- Zbudować messaging marketingowy na podstawie tych danych
- Cel: zainteresować decydentów (C-level, Head of GPS)

### 4.3 Konsolidacja dotychczasowych rozmów i wniosków strategicznych

- Zebrać wszystkie transkrypcje ze spotkań (w tym dzisiejsze)
- Wrzucić do projektu w AI (Perplexity / innym narzędziu)
- Zasilić to: notatkami, transkrypcjami, wynikami ankiet, zdjęciem tablicy ze spotkania z Tomem
- Cel: mieć jedno źródło wiedzy, z którego AI może wygenerować spójny opis produktu / oferty

---

## 5. Projekt Practice Lab — szczegóły konstrukcji

### Struktura produktowa

- **Basic** — pozostaje bez zmian
- **Practice Lab** (zastępuje "Intermediate") — dostosowane ćwiczenia praktyczne po Basic'u
- **Hackathon / Workshop** — format warsztatowy, min. 1 dzień, docelowo 3–4 dni przy poważnych tematach

### Pre-requisyt: kwestionariusz uczestników

Oferta ma zawierać jasną informację: **bez wypełnionego kwestionariusza nie przeprowadzimy sensownego Practice Lab.**

Kwestionariusz wysyłany 3–4 tygodnie przed, zbiera:
- Dostępne narzędzia technologiczne (Python, licencja Premium, Copilot Studio, Power Automate)
- Obszary procesowe, w których pracują uczestnicy
- Poziom zaawansowania
- Oczekiwania i braki po Basic'u

### Ścieżki Practice Lab w zależności od odpowiedzi

| Konfiguracja firmy | Zawartość Practice Lab |
|---|---|
| Tylko Copilot Chat (bez Pythona, bez Premium) | Ćwiczenia z asystentami tekstowymi, promptowanie zaawansowane |
| Premium, bez Pythona | Power Automate workflows, zaawansowane asystenty, Copilot Studio |
| Python dostępny | PDF processing, analiza danych, budowanie własnych agentów, mini-apki HTML |

### Repozytorium ćwiczeń

Plan stworzenia zbioru **20–50 micro-ćwiczeń**, m.in.:
- HTML: dashboardy, wykresy, harmonogramy Gantta, strony prezentacyjne
- HTML: dokumentacja projektowa (one-page scrollable)
- Micro-apki: raportowanie obecności w biurze (HR), inne narzędzia operacyjne
- Copilot Studio: agenci do zarządzania procesami
- Six Sigma: storyboard dokumentujący projekt (historycznie robiony w PowerPoint — teraz przez AI)
- Agenci do przetwarzania maili i Teamsów z telefonu (mowa o use case "pierdolę maile z auta")

Zasada: kwestionariusz przed praktyce labem **automatycznie segreguje uczestników** na odpowiednie ćwiczenia — trener ma gotową "playlistę" bez gry w zgadywanie.

---

## 6. Automatyzacja tworzenia szkoleń — missed lead

Wzmianka o leadzie ze Strajkera (nie doszło do współpracy): firma szukała kogoś, kto pomoże **zautomatyzować produkcję treści szkoleniowych**.

Koncept:
- Wrzucasz PowerPoint → AI generuje z tego wideo z narratorem + przerywniki ćwiczeniowe
- Skierowane do dużych firm chcących produkować e-learningi na skalę

Wniosek: temat warto mieć w tyle głowy jako potencjalny produkt lub komponent oferty.

---

## 7. Konferencja w Lizbonie — taktyka i narzędzia

**Termin:** 27–28 maja

### Ankieta "Co byś zrobił z AI?"

Pomysł: zbudować prostą ankietę (1–5 pytań) i wdrożyć ją jako aplikację dostępną z telefonu.

- Na konferencji: w rozmowie z kimkolwiek, prosisz go o wypełnienie ankiety na jego telefonie
- Zbierasz dane rynkowe na żywo
- Demo "produktu AI" — apka sama w sobie jest przykładem możliwości
- Idealna do uruchomienia na serwerze testowym (gotowe środowisko)

### Cel networkingowy

- Mieć 2–3 przygotowane "user stories" / opisy rozwiązań do opowiadania
- Umawia się kol **od razu na miejscu** (nie po konferencji — wtedy tematy giną)
- Materiały do wysyłki po spotkaniu: coś zwięzłego, zrozumiałego, z klarownym CTA

---

## 8. Webinary — plan działania

- **Czas:** sierpień (sezon ogórkowy = dobry czas na treści online)
- **Format:** 2 webinary jako rozgrzewka przed falą sprzedaży od września
- **Potencjalny temat:** jak zarządzać procesami z myślą o AI, zaproszeni goście (np. Piotr Pieżak z Nowardisku — GPO jako product ownerzy rozwiązań AI)
- **Cel:** brand building + generowanie leadów na Q4

---

## 9. Akcje do realizacji (podsumowanie)

| # | Akcja | Odpowiedzialny |
|---|---|---|
| 1 | Aktualizacja prezentacji sprzedażowej + zmiana nazwy na "Practice Lab" | Robert / Paweł |
| 2 | Mailing: ostatnie miejsca przed wakacjami na szkolenia AI | Edyta (dane → Robert) |
| 3 | Przegląd ankiet Slido → wyciągnięcie braków → agenda Practice Lab | Robert |
| 4 | Kwestionariusz pre-requisyt do Practice Lab (build) | Robert / Paweł |
| 5 | Repozytorium ćwiczeń — szkic z istniejących pomysłów | Robert + Szymon |
| 6 | Wrzucenie transkrypcji + notatek do projektu AI (Perplexity) | Robert |
| 7 | Sprawdzenie grup po freelancera szkoleniowego | Paweł |
| 8 | Ankieta "Co byś zrobił z AI?" — budowa apki na Lizbonę | Robert |
| 9 | Zaplanowanie 2 webinarów na sierpień | Robert / Krystian |
| 10 | Spotkanie z Tomem + zdjęcie tablicy → do projektu AI | Robert |
