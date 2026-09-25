"""Database connection helper for the Advance Tracker Client report app."""

import os
from urllib.parse import quote_plus

import pandas as pd
from sqlalchemy import bindparam, create_engine, text


def get_engine():
    host = os.getenv("DB_HOST")
    port = os.getenv("DB_PORT", "3306")
    name = os.getenv("DB_NAME")
    user = os.getenv("DB_USER")
    password = quote_plus(os.getenv("DB_PASSWORD", ""))

    url = f"mysql+pymysql://{user}:{password}@{host}:{port}/{name}"
    return create_engine(url)


def run_query(query: str) -> pd.DataFrame:
    engine = get_engine()
    try:
        with engine.connect() as conn:
            return pd.read_sql(query, conn)
    finally:
        engine.dispose()


def run_query_params(query: str, params: dict | None = None) -> pd.DataFrame:
    """Run a parameterized query. List/tuple/set values are bound as expanding
    IN-clause parameters automatically."""
    params = params or {}
    stmt = text(query)
    for key, value in params.items():
        if isinstance(value, (list, tuple, set)):
            stmt = stmt.bindparams(bindparam(key, expanding=True))

    engine = get_engine()
    try:
        with engine.connect() as conn:
            return pd.read_sql(stmt, conn, params=params)
    finally:
        engine.dispose()
