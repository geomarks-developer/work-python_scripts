# saved_queries.py

import sqlite3
from datetime import datetime
import streamlit as st

DB_FILE = "bi_tool.db"


def init_db():

    conn = sqlite3.connect(DB_FILE)

    conn.execute("""
        create table if not exists query_history (
            id integer primary key autoincrement,
            query_text text,
            run_ts text
        )
    """)

    conn.execute("""
        create table if not exists saved_queries (
            id integer primary key autoincrement,
            query_name text,
            query_text text,
            created_ts text
        )
    """)

    conn.commit()
    conn.close()


def add_history(sql):

    conn = sqlite3.connect(DB_FILE)

    conn.execute(
        """
        insert into query_history
        (query_text, run_ts)
        values (?,?)
        """,
        (
            sql,
            datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        ),
    )

    conn.commit()
    conn.close()


def get_history(limit=20):

    conn = sqlite3.connect(DB_FILE)

    rows = conn.execute(
        """
        select id,
               query_text,
               run_ts
        from query_history
        order by id desc
        limit ?
        """,
        (limit,),
    ).fetchall()

    conn.close()

    return rows


def save_query(name, sql):

    user_id = (
        st.session_state.get(
            "oracle_user_input",
            ""
        )
        .strip()
        .upper()
    )

    current_ts = datetime.now().strftime(
        "%Y-%m-%d %H:%M:%S"
    )

    conn = sqlite3.connect(DB_FILE)

    conn.execute(
        """
        INSERT INTO saved_queries
        (
            query_name,
            query_text,
            user_id,
            created_ts,
            updated_ts
        )
        VALUES
        (
            ?, ?, ?, ?, ?
        )
        ON CONFLICT (
            upper(user_id),
            upper(query_name)
        )
        DO UPDATE SET
            query_text = excluded.query_text,
            updated_ts = excluded.updated_ts
        """,
        (
            name,
            sql,
            user_id,
            current_ts,
            current_ts,
        ),
    )

    conn.commit()
    conn.close()


def get_saved_queries():

    user_id = (
        st.session_state.get(
            "oracle_user_input",
            ""
        )
        .strip()
        .upper()
    )

    conn = sqlite3.connect(DB_FILE)

    rows = conn.execute(
        """
        select id,
               query_name,
               query_text
        from saved_queries
        where upper(user_id) = ?
        order by query_name
        """,
        (user_id,)
    ).fetchall()

    conn.close()

    return rows


def delete_saved_query(query_id):

    conn = sqlite3.connect(DB_FILE)

    conn.execute(
        """
        delete from saved_queries
        where id = ?
        """,
        (query_id,),
    )

    conn.commit()
    conn.close()

# --------------------------------------------------
# Helpers
# --------------------------------------------------

def get_favorite_objects():

    conn = sqlite3.connect(DB_FILE)

    rows = conn.execute(
        """
        select object_name,
               object_type,
               category
        from favorite_objects
        order by display_order
        """
    ).fetchall()

    conn.close()

    return rows