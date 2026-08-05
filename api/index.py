import os
import sys

# Ensure parent directory is in sys.path so app.py can be imported cleanly in Vercel
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from app import app
