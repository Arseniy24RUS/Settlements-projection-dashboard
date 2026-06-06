# Contributing

Thank you for improving `Settlements-projection-dashboard`. The project welcomes small fixes, documentation improvements, tests, and research-methodology questions.

## Working Locally

```bash
npm install
npm test
```

For manual inspection:

```bash
npx http-server . -a 127.0.0.1 -p 4175 -c-1
```

Open <http://127.0.0.1:4175/>.

## Contribution Rules

- Keep pull requests focused on one change.
- Do not mix data refreshes, UI changes, and methodology edits unless the issue explicitly requires it.
- Preserve bilingual EN/RU user-facing documentation where practical.
- For data or methodology changes, update `docs/methodology.md` and describe provenance, assumptions, and affected files.
- For interface text changes, update `locales/en.json`, `locales/ru.json`, and tests as needed.
- For visual changes, include screenshots or describe what was checked manually.
- Respect `LICENSE`, `LICENSE-DOCS-AND-DATA.md`, and `THIRD_PARTY_NOTICES.md`.

## Pull Request Checklist

- The change has a clear issue, motivation, or research question.
- Local tests pass or the reason for not running them is documented.
- New or changed data have source, version, and preprocessing notes.
- README changes are minimal and linked to deeper docs where possible.
- Third-party data, library, or map-service implications are documented.

---

# Участие В Проекте

Спасибо за помощь в развитии `Settlements-projection-dashboard`. Проект принимает небольшие исправления, улучшения документации, тесты и методологические вопросы.

## Локальная Работа

```bash
npm install
npm test
```

Для ручной проверки:

```bash
npx http-server . -a 127.0.0.1 -p 4175 -c-1
```

Откройте <http://127.0.0.1:4175/>.

## Правила Вклада

- Делайте pull request сфокусированным на одном изменении.
- Не смешивайте обновление данных, изменение интерфейса и методологические правки без явной необходимости.
- По возможности сохраняйте двуязычную EN/RU документацию.
- При изменении данных или методики обновляйте `docs/methodology.md` и описывайте происхождение, предпосылки и затронутые файлы.
- При изменении текстов интерфейса обновляйте `locales/en.json`, `locales/ru.json` и тесты при необходимости.
- При визуальных изменениях прикладывайте скриншоты или описывайте ручную проверку.
- Соблюдайте `LICENSE`, `LICENSE-DOCS-AND-DATA.md` и `THIRD_PARTY_NOTICES.md`.
