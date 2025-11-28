import sys
import os

# Tambahkan direktori project ke sys.path
sys.path.insert(0, os.path.dirname(__file__))

# Import aplikasi Flask
from app import app as application
