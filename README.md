# SIFET Estatales

Sistema de Información Financiera de Entes Estatales - Órgano de Fiscalización Superior del Estado de Tlaxcala

**Status:** Active development

## Description

REST API service providing financial information system for state entities. Features health check endpoints and JSON API responses.

## Tech Stack

- **Backend:** Flask 3.0.3
- **Server:** Gunicorn
- **Python:** 3.8+

## Quick Start

```bash
# Setup virtual environment
python3 -m venv venv
source venv/bin/activate

# Install dependencies
pip install -r requirements.txt

# Run application
python run.py
```

## API Endpoints

- `GET /` - Service information
- `GET /health` - Health check endpoint

## Configuration

Uses environment variables:
- `SECRET_KEY` - Flask session secret (default: auto-generated)
- `PORT` - Server port (default: 5008)
