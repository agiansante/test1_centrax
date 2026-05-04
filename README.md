# LLM Chat Test

Prima applicazione di test con una chat web e un backend Node che chiama la OpenAI Responses API.

## Requisiti

- Node.js 18 o superiore
- Una API key OpenAI

## Avvio

Su PowerShell:

```powershell
$env:OPENAI_API_KEY="la-tua-api-key"
npm start
```

Poi apri:

```text
http://localhost:3000
```

Puoi cambiare modello con:

```powershell
$env:OPENAI_MODEL="gpt-5.2"
```

## File principali

- `server.js`: server HTTP e endpoint `/api/chat`
- `public/index.html`: interfaccia chat
- `public/styles.css`: stile della pagina
- `public/app.js`: logica frontend della chat
