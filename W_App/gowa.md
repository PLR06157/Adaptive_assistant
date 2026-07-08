# GOWA – ściągawka z komendami (go-whatsapp-web-multidevice)

Narzędzie: [aldinokemal/go-whatsapp-web-multidevice](https://github.com/aldinokemal/go-whatsapp-web-multidevice)

W przeciwieństwie do `whatsapp-cli` – GOWA to serwer (REST API + web UI + tryb MCP), nie narzędzie z bezpośrednimi komendami typu `send`. Do wysyłania/odczytywania wiadomości trzeba wołać jego REST API.

## Instalacja

### Docker (zalecane)

```bash
docker run --detach --publish=3000:3000 --name=whatsapp \
  --restart=always --volume=$(docker volume create --name=whatsapp):/app/storages \
  aldinokemal2104/go-whatsapp-web-multidevice rest
```

Docker Compose:

```yaml
services:
  whatsapp:
    image: aldinokemal2104/go-whatsapp-web-multidevice
    restart: always
    ports: ["3000:3000"]
    volumes: [whatsapp:/app/storages]
    command: [rest, --basic-auth=admin:admin, --port=3000]
volumes:
  whatsapp:
```

### Binarka / źródła

Gotowe binarki (Linux/macOS/Windows/ARM) do pobrania z [wydań na GitHubie](https://github.com/aldinokemal/go-whatsapp-web-multidevice/releases) albo kompilacja ze źródeł (wymaga Go 1.25.5+).

## Uruchomienie

Dwa tryby pracy binarki:

```bash
./GOWA rest    # serwer REST API + web UI
./GOWA mcp     # serwer MCP (do podpięcia jako narzędzie w Claude Code / agentach)
```

Pełna lista flag: `./GOWA --help`.

## Logowanie (QR code)

1. Wejdź w przeglądarce na `http://localhost:3000` – strona pokaże kod QR.
2. Zeskanuj w telefonie: WhatsApp → Ustawienia → Urządzenia powiązane → Połącz urządzenie.

Kod QR jest też dostępny bezpośrednio pod `GET /app/login`.

## Konfiguracja (najważniejsze flagi/env)

| Flaga | Env var | Opis |
|---|---|---|
| `--port` | `APP_PORT` | port serwera (domyślnie `3000`) |
| `--basic-auth=user:pass` | `APP_BASIC_AUTH` | podstawowa autoryzacja HTTP |
| `--webhook=URL` | `WHATSAPP_WEBHOOK` | adres webhooka do zdarzeń |
| `--webhook-secret=SECRET` | `WHATSAPP_WEBHOOK_SECRET` | sekret do weryfikacji webhooka |
| `--webhook-ignore-jids=...` | – | filtr JID-ów ignorowanych w webhookach (`@g.us`, `@s.whatsapp.net`, `@lid`) |
| – | `CHATWOOT_DAYS_LIMIT_IMPORT_MESSAGES` | limit dni historii importowanej do Chatwoota |

Konfiguracja też przez plik `.env`.

## Czaty (rozmowy)

```bash
curl "http://localhost:3000/chats?limit=25"
curl "http://localhost:3000/chats?search=Jan"
curl "http://localhost:3000/chats?archived=true"
```

Parametry: `limit` (domyślnie 25, max 100), `offset`, `search`, `has_media`, `archived`.

## Wiadomości

```bash
curl "http://localhost:3000/chat/48661662016@s.whatsapp.net/messages?limit=20"
curl "http://localhost:3000/chat/120363427408489520@g.us/messages?limit=50&offset=0"
```

Parametry: `limit` (domyślnie 50, max 100), `offset`, `start_time` / `end_time` (ISO 8601), `is_from_me`, `media_only`, `search` (szukanie po treści w danym czacie).

## Wysyłanie wiadomości

```bash
curl -X POST "http://localhost:3000/send/message" \
  -H "Content-Type: application/json" \
  -d '{"phone": "120363427408489520@g.us", "message": "Próba mikrofonu 🎙️"}'
```

Do grupy – w polu `phone` podać JID kończący się na `@g.us`. Przy wielu urządzeniach dodać nagłówek `X-Device-Id` lub parametr `?device_id=`.

## Media

```bash
curl "http://localhost:3000/message/{message_id}/download" -o plik
```

## Webhooki

- Globalne: flagi/env z tabeli wyżej.
- Per-urządzenie: `PATCH /devices/:device_id/webhook`.
- Filtrowanie zdarzeń: `WHATSAPP_WEBHOOK_EVENTS`.
- Ignorowanie JID-ów: `--webhook-ignore-jids` (obsługuje `@g.us`, `@s.whatsapp.net`, `@lid`).

## Tryb MCP (integracja z Claude Code)

```bash
./whatsapp mcp
```

Serwer wystawia transport SSE, domyślnie pod `http://localhost:8080/sse`. Konfiguracja klienta MCP:

```json
{
  "mcpServers": {
    "whatsapp": {
      "url": "http://localhost:8080/sse"
    }
  }
}
```

## Format JID

- Czat prywatny: `numer@s.whatsapp.net`
- Grupa: `id@g.us`
- Kontakt zidentyfikowany przez LID (nowy format WhatsApp): `id@lid`

## Pełna dokumentacja API

Kompletna specyfikacja endpointów, parametrów i schematów odpowiedzi: `docs/openapi.yaml` w repozytorium, do podejrzenia np. przez SwaggerEditor.
