"""
Configuration for SIFET Estatales
"""
import os


class Config:
    """Base configuration"""
    SECRET_KEY = os.getenv('SECRET_KEY', 'sifet-estatales-secret-key')
    PORT = int(os.getenv('PORT', '5008'))
