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
- Jeśli chcesz trzymać CSV lub szablon w innym miejscu, podaj nowe ścieżki.

## 3. Plik CSV z adresami

Domyślnie skrypt czyta `mailing/recipients.csv`. Struktura:

```csv
email,first_name,sender_name,subject
jan.kowalski@example.com,Jan,Adaptive Group,"Oferta szkoleniowa Adaptive Group"
```

- `email` – adres odbiorcy (możesz zmienić nazwę kolumny przez `EMAIL_COLUMN`).
- `subject` – tytuł wiadomości; jeśli puste możesz ustawić `DEFAULT_SUBJECT`.
- Pozostałe kolumny są wykorzystywane do wypełnienia pól w szablonie HTML.

## 4. Szablon HTML

Plik `mailing/email_template.html` wykorzystuje formatowanie `str.format`, więc każdy nawias `{pole}` musi mieć odpowiadającą kolumnę w CSV (np. `{first_name}`). Możesz dodać własne placeholdery.

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

- `--csv`, `--template`, `--attachment`
- `--email-column`, `--subject-column`, `--default-subject`
- `--log-level` (`INFO`, `DEBUG` itd.)

## 6. Debugowanie

- Komunikat „Unable to acquire access token” oznacza problem z uprawnieniami aplikacji lub błędne dane `.env`.
- Jeśli wiersz w CSV nie ma adresu e-mail, zostaje pominięty (log na poziomie `WARNING`).
- Użyj `--log-level DEBUG`, aby zobaczyć więcej szczegółów.

## 7. Dalsze kroki

- Dodaj kolumny do CSV, aby personalizować treść (np. `company_name`, `meeting_link`).
- Połącz z narzędziem harmonogramu (cron, GitHub Actions) jeśli chcesz wysyłać cyklicznie.
