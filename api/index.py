"""Vercel serverless entrypoint.

Vercel's @vercel/python runtime imports this module and serves the WSGI
callable named ``app``. We expose the Flask app defined in the project root.
"""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from app import app  # noqa: E402  (Flask WSGI app, served by Vercel)
