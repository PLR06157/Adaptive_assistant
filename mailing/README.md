# Automatyczna wysyłka mailingowa (Microsoft 365)

Poniższe narzędzie pozwala wysyłać spersonalizowane wiadomości HTML z załącznikiem z firmowego konta Microsoft 365 (Exchange Online), korzystając z Microsoft Graph API.

## 1. Rejestracja aplikacji w Azure AD

1. Zaloguj się do [Azure Portal](https://portal.azure.com) kontem administratora M365.
2. Przejdź do **Azure Active Directory → App registrations → New registration**.
3. Nadaj nazwę, pozostaw typ konta „Accounts in this organizational directory only”.
4. Po utworzeniu aplikacji:
   - Zanotuj **Application (client) ID** oraz **Directory (tenant) ID**.
   - W sekcji **Certificates & secrets** dodaj **Client secret** – zapisz jego wartość, nie będzie dostępna ponownie.
5. W sekcji **API permissions** dodaj:
   - `Mail.Send` (Application).
   - zaznacz **Grant admin consent**.

## 2. Przygotowanie środowiska lokalnego

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r mailing/requirements.txt
```

Skopiuj plik `mailing/.env.example` do `mailing/.env` i wypełnij:

- `TENANT_ID`, `CLIENT_ID`, `CLIENT_SECRET` – dane z Azure AD.
- `SENDER_EMAIL` – adres konta, z którego wysyłasz wiadomości.
- `ATTACHMENT_PATH` – ścieżka do pliku PDF/DOCX itp., który ma być dołączony.
- `XLSX_PATH`, `RECIPIENT_SHEET_NAME` – opcjonalnie wskaż inną lokalizację pliku i arkusz z danymi.
- `SAVE_TO_SENT_ITEMS=false` – ustaw, jeśli nie chcesz zapisywać wysłanych wiadomości w skrzynce nadawczej.

## 3. Plik XLSX z adresami

Domyślnie skrypt czyta `mailing/recipients.xlsx` (aktywny arkusz albo nazwany przez `RECIPIENT_SHEET_NAME`). Kolumny A–D muszą występować w kolejności:

```
A: email
B: first_name
C: sender_name
D: subject
```

Nagłówek w pierwszym wierszu jest opcjonalny, ale zalecany (skrypt sam go wykryje). W kolumnie `subject` możesz pozostawić puste pola i zdefiniować `DEFAULT_SUBJECT` w `.env` lub linii poleceń. Pozostałe kolumny (`first_name`, `sender_name`) trafiają do kontekstu szablonu HTML. Jeśli potrzebujesz dodatkowych placeholderów, dodaj je jako nowy arkuszowy blok i zmodyfikuj szablon.

## 4. Szablon HTML

Plik `mailing/email_template.html` wykorzystuje formatowanie `str.format`, więc każdy nawias `{pole}` musi mieć odpowiadającą wartość w arkuszu (np. `{first_name}`). Możesz dodać własne placeholdery i uzupełnić je dodatkowymi kolumnami.

## 5. Wysyłka

Przed pierwszym uruchomieniem warto zrobić „suchą próbę”:

```bash
python mailing/send_mail.py --dry-run
```

Realna wysyłka:

```bash
python mailing/send_mail.py
```

Parametry, które możesz nadpisać z linii poleceń:

- `--xlsx`, `--sheet-name`, `--template`, `--attachment`
- `--default-subject`, `--min-wait`, `--max-wait`
- `--log-level` (`INFO`, `DEBUG` itd.), `--dry-run`
- `--no-save-to-sent-items` aby nie archiwizować wysyłek w folderze „Wysłane”

## 6. Debugowanie

- Komunikat „Unable to acquire access token” oznacza problem z uprawnieniami aplikacji lub błędne dane `.env`.
- Jeśli wiersz w arkuszu nie ma adresu e-mail, zostaje pominięty (log na poziomie `WARNING`).
- Użyj `--log-level DEBUG`, aby zobaczyć więcej szczegółów.

## 7. Dalsze kroki

- Dodaj kolumny do arkusza, aby personalizować treść (np. `company_name`, `meeting_link`).
- Połącz z narzędziem harmonogramu (cron, GitHub Actions) jeśli chcesz wysyłać cyklicznie.
