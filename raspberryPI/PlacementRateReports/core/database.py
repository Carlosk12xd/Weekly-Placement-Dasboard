"""Load live database dependencies/settings only when a database build is requested."""
import os
from pathlib import Path


def load_environment(directory):
    """Read one report's settings, without leaking them into the next report.

    Explicit process settings override that directory's .env, as before. Reading
    a file never sets process variables. Callers must use the returned mapping.
    """
    path = Path(directory) / '.env'
    settings = {}
    if path.is_file():
        try:
            from dotenv import dotenv_values
        except ImportError as exc:
            raise RuntimeError('Install requirements.txt to read .env settings') from exc
        settings = {k: v for k, v in dotenv_values(path).items() if v is not None}
    return {**settings, **os.environ}


def connect_database(directory):
    settings = load_environment(directory)
    required = ('DB_HOST', 'DB_USER', 'DB_PASSWORD', 'DB_NAME')
    missing = [key for key in required if not settings.get(key)]
    if missing:
        raise RuntimeError('Missing database settings: ' + ', '.join(missing))
    import mysql.connector
    return mysql.connector.connect(host=settings['DB_HOST'], user=settings['DB_USER'],
        password=settings['DB_PASSWORD'], database=settings['DB_NAME'], autocommit=False)
