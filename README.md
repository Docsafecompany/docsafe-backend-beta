# DocSafe Backend (beta)


Endpoints:
- `GET /health`
- `POST /clean` (form-data: file, strictPdf = 'true'|'false')
- `POST /clean-v2` (form-data: file, strictPdf = 'true'|'false')


Réponses: `application/zip` contenant `cleaned.<ext>` + `report.html`.


Env:
- `LT_ENDPOINT`, `LT_LANGUAGE` (optionnel)
- `AI_PROVIDER`, `OPENAI_API_KEY`, `AI_MODEL`, `AI_TEMPERATURE`
- `CORS_ORIGIN`, `PORT`
