# Contexto del proyecto — SIFET Estatales

## Descripción
Sistema de Fiscalización de Entes Estatales del OFS Tlaxcala. Gestiona el historial de titulares y administrativos de entes estatales. Interfaz diferenciada por rol: `viewer` (luis) ve datos, `loader` (gabo) carga Excel.

## Usuarios
- **luis** (`viewer`) — visualización y consulta
- **gabo** (`loader`) — carga de archivos, administración

## Base de datos
`sifeet.db` — SQLite local. Backups en `backups/`.

## Auth
`USERS` dict construido desde `USER_CREDENTIALS` env var: `usuario:contraseña:rol,...`  
`login_required` / `luis_required` / `gabo_required` decorators.

## Estado de migración
- Migrado en wave 3 (2026-04-13)
- Contraseñas hardcodeadas → `USER_CREDENTIALS` env var
- `desploy/` → `deploy/`
- `/api/health` añadido
- `AGENTS.md` y `CLAUDE.md` movidos a `prompts/`
