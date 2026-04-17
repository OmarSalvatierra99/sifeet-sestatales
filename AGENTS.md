# Repository Guidelines

## Project Structure & Module Organization
`app.py` is the Flask entrypoint: it initializes config, auth, database access, and shared helpers. Route logic is split by role in `scripts/gabo_routes.py` and `scripts/luis_routes.py`. Jinja templates live in `templates/`; frontend assets are under `static/css`, `static/js`, `static/img`, and `static/vendor`. Tests live in `tests/` with shared setup in `tests/conftest.py`. Operational files include `sifeet.db`, `backups/db/`, `logs/`, and deployment templates in `deploy/nginx` and `deploy/systemd`.

## Build, Test, and Development Commands
Create an environment and install dependencies with `python -m venv venv && source venv/bin/activate && pip install -r requirements.txt`. Start the app locally with `python app.py`; the current default is `http://127.0.0.1:5008`. Run tests with `pytest`. For quick health verification, use `curl http://127.0.0.1:5008/api/health`.

## Coding Style & Naming Conventions
Follow existing Python style: 4-space indentation, `snake_case` for functions/variables, and `UPPER_CASE` for module constants such as `DB_PATH` or `UID_PREFIX`. Keep Flask routes thin when possible and move reusable logic into helper functions near the owning module. Template names should stay descriptive and feature-based, for example `carga_observaciones_admin.html`. Preserve current Spanish domain wording in UI labels and business rules.

## Testing Guidelines
Tests use `pytest` and Flask’s test client. Add new tests in `tests/test_app.py` or a new `tests/test_<feature>.py` module. Name tests by behavior, for example `test_login_with_valid_credentials`. Cover route status codes, auth behavior, and any data normalization you change. No coverage gate is configured, so treat regression coverage as part of the change.

## Commit & Pull Request Guidelines
Recent history uses short progress-style messages such as `Working` and dated Spanish summaries. Prefer clearer imperative subjects that name the area changed, for example `Normalize period parsing for carga masiva`. PRs should include: a short description, affected routes/templates, any database or `.env` changes, test results, and screenshots for template or CSS updates.

## Security & Configuration Tips
Copy `.env.example` to `.env` and set a real `SECRET_KEY` for non-test use. Do not hardcode credentials; shared users come from the external catalog referenced in `.env.example`. Treat `sifeet.db` and backup files as sensitive operational data, and avoid manual edits outside the app or migration scripts.
