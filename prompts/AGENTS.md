# Repository Guidelines

## Project Structure & Module Organization
`app.py` is the main Flask entrypoint and initializes the SQLite schema in `sifeet.db`. Shared backup helpers live in `backup_utils.py`. Route-heavy modules are split into [`scripts/gabo_routes.py`](/home/gabo/portfolio/projects/07-sifet-estatales/scripts/gabo_routes.py) and [`scripts/luis_routes.py`](/home/gabo/portfolio/projects/07-sifet-estatales/scripts/luis_routes.py), while one-off maintenance and import utilities also live under `scripts/`. HTML templates are in `templates/`, frontend assets are in `static/css`, `static/js`, and `static/img`, and production service files are in `deploy/` (systemd, nginx, env).

## Build, Test, and Development Commands
Use the local virtualenv when possible.

- `python3 -m venv venv && source venv/bin/activate && pip install -r requirements.txt`: create a local environment and install Flask plus `openpyxl`.
- `python app.py`: start the development server on `0.0.0.0:5008`.
- `python scripts/prune_backups.py --show 5`: preview old backup cleanup without deleting files.
- `python scripts/prune_backups.py --apply`: delete backups according to the script defaults.
- `python -m py_compile app.py backup_utils.py scripts/*.py`: quick syntax smoke test for Python changes.

## Coding Style & Naming Conventions
Follow current Python style: 4-space indentation, `snake_case` for functions and variables, `UPPER_SNAKE_CASE` for constants, and small helper functions for repeated normalization or parsing logic. Keep route handlers thin where possible and move reusable database or backup logic into helpers. Match the existing naming in templates and forms; many domain terms are Spanish and should stay consistent with the UI and database columns.

## Testing Guidelines
There is no automated test suite in the repository yet. For every change, run the `py_compile` smoke check and manually verify the affected screen or script against a safe copy of `sifeet.db`. For import or cleanup scripts, test against a backup first and confirm both success messages and resulting database state.

## Commit & Pull Request Guidelines
Recent history uses short, informal subjects, but new commits should be clearer and imperative, for example `Add backup pruning dry-run summary`. Keep commits focused on one concern. Pull requests should summarize user-visible changes, note any database or script side effects, include screenshots for template or CSS updates, and list the manual verification steps you ran.

## Security & Configuration Tips
Do not rely on the fallback `SECRET_KEY` outside local development. Treat `sifeet.db` and `backups/` as sensitive data, avoid committing generated backups, and create a fresh backup before running any script that updates or deletes records.
