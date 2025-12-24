"""
Entry point for SIFET Estatales application
"""
from app import create_app
from app.config import Config

app = create_app()

if __name__ == "__main__":
    app.logger.info(f"Starting SIFET Estatales on port {Config.PORT}")
    app.run(host="0.0.0.0", port=Config.PORT, debug=False)
