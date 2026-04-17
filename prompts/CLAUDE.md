# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Overview

SIFEET Estatales is a Flask web app for tracking fiscal year audit observations for state entities. It uses SQLite as its database and is deployed at `sifet-estatales.omar-xyz.shop`.

## Running the app

```bash
# Activate the virtual environment first
source venv/bin/activate

# Run the dev server (port 5008)
python app.py

# With template auto-reload
TEMPLATES_AUTO_RELOAD=1 python app.py
```

There are no tests. There is no linter configured.

## Architecture

**Entry point:** `app.py` — defines the Flask app, all shared utility functions, DB schema, and the `ROUTE_DEPS` dict. It calls `init_db()` at module load time and registers routes from both blueprint modules.

**Route modules** (in `scripts/`):
- `scripts/gabo_routes.py` — routes under `/carga/*` for the `gabo` (loader) user: data entry, bulk upload of observations via CSV, titulares management.
- `scripts/luis_routes.py` — routes for the `luis` (viewer) user: dashboard, filtering, Excel exports, comparativo anual, stats.

Both modules use a dependency injection pattern: `register_gabo_routes(app, deps)` / `register_luis_routes(app, deps)` receive a `ROUTE_DEPS` dict from `app.py` and call `globals().update(deps)` to import shared helpers into their scope.

**Database:** `sifeet.db` (SQLite, WAL mode). Schema is created by `init_db()` in `app.py`. Key tables:
- `observaciones` — audit observations per entity/exercise
- `entes_detalle` — entity master data with `ente_uid` cross-year identifier (`ENT-N` format)
- `historial_titulares` — responsible officers per entity/period
- `registros` — legacy records table
- `fuentes_financiamiento`, `catalogo_irregularidades` — lookup tables

**Auth:** Two hardcoded users in `USERS` dict in `app.py`. `luis` = viewer role, `gabo` = loader role. Decorators: `@luis_required`, `@gabo_required`, `@login_required`.

**Backups:** `backup_utils.py` handles SQLite hot backups to `backups/` directory. `DB_SNAPSHOT_KEEP_COUNT = 30` in `gabo_routes.py` controls retention. `scripts/prune_backups.py` is a standalone script.

**Templates:** Jinja2 in `templates/`. No JS framework — vanilla JS with inline `<script>` tags. CSS in `static/css/style.css`.

## Key conventions

- `ente_uid` (`ENT-N`) is a stable cross-year entity identifier. `ente_id` is the per-year numeric ID from the source data. Use `normalize_ente_id_sql()` when comparing `ente_id` in SQL (strips trailing dots/spaces).
- Period strings use Spanish format: `"01 de enero al 31 de diciembre"`. `parse_periodo_cedula()` converts these to `YYYY-MM-DD` date pairs.
- `normalize_text_key()` strips accents and lowercases for fuzzy matching.
- Ejercicio (fiscal year) is stored as a string (e.g. `"2025"`). `GABO_READONLY_EJERCICIOS = {"2023", "2024"}` prevents edits to past years.
- Amounts in the DB are floats. Input parsing strips commas before `float()`.
