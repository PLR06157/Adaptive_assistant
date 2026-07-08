# WhatsApp CLI – ściągawka z komendami

Narzędzie: [vicentereig/whatsapp-cli](https://github.com/vicentereig/whatsapp-cli)

## Instalacja (macOS)

```bash
brew install vicentereig/tap/whatsapp-cli
whatsapp-cli version
```

## Autoryzacja

```bash
whatsapp-cli auth
```

Zeskanuj kod QR w WhatsApp na telefonie: Ustawienia → Urządzenia powiązane → Połącz urządzenie.
Sesja utrzymuje się bez ponownego skanowania przez ok. 20 dni.

## Synchronizacja

Musi działać w osobnym terminalu (lub w tle) przez cały czas, żeby CLI odbierało wiadomości i pobierało historię:

```bash
whatsapp-cli sync
```

- Ilość zsynchronizowanej historii zależy od retencji po stronie serwerów WhatsApp – to nie jest błąd narzędzia.
- Jeśli połączenie się zerwie (np. błąd websocket/EOF), po prostu uruchom `sync` ponownie.
- Jeśli sesja zniknie z listy „Urządzenia powiązane" na telefonie, trzeba przejść przez `auth` od nowa.

## Czaty (rozmowy)

```bash
whatsapp-cli chats list
whatsapp-cli chats list | jq
```

## Wiadomości

Lista wiadomości z konkretnego czatu (JID bierzesz z `chats list`):

```bash
whatsapp-cli messages list --chat 135137455042659@lid
whatsapp-cli messages list --chat NUMER@s.whatsapp.net
whatsapp-cli messages list --chat NUMER@s.whatsapp.net --limit 100 --page 1
```

Wyszukiwanie wiadomości po treści:

```bash
whatsapp-cli messages search --query "szukana fraza"
whatsapp-cli messages search --query "szukana fraza" --limit 100
```

## Kontakty

```bash
whatsapp-cli contacts search --query "Jan Kowalski"
```

## Wysyłanie wiadomości

Do osoby prywatnej (numer w formacie międzynarodowym, bez `+` i spacji):

```bash
whatsapp-cli send --to 48123456789 --message "Treść wiadomości"
```

Do grupy (JID kończący się na `@g.us`, weź z `chats list`):

```bash
whatsapp-cli send --to 123456789@g.us --message "Treść wiadomości"
```

## Media

```bash
whatsapp-cli media download
```

## Format JID

- Czat prywatny: `numer@s.whatsapp.net`
- Grupa: `id@g.us`

## Format odpowiedzi

Wszystkie komendy zwracają JSON:

```json
{"success": true, "data": ..., "error": null}
```
