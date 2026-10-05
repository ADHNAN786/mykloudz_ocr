import os
import sqlite3
import hashlib
import secrets
from typing import List, Dict, Any, Tuple, Optional
import pandas as pd

DB_PATH = os.path.join(os.path.dirname(__file__), "auth_usage.db")


def get_connection():
    return sqlite3.connect(DB_PATH)


def hash_password(password: str, salt: str) -> str:
    """Generate SHA-256 salted hash."""
    return hashlib.sha256((salt + password).encode("utf-8")).hexdigest()


def init_db():
    """Initialize database tables and seed default admin user if not present."""
    conn = get_connection()
    cursor = conn.cursor()
    
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS users (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            email TEXT UNIQUE NOT NULL COLLATE NOCASE,
            password_hash TEXT NOT NULL,
            salt TEXT NOT NULL,
            role TEXT NOT NULL DEFAULT 'user',
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    """)
    
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS token_usage (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            user_email TEXT NOT NULL COLLATE NOCASE,
            file_name TEXT,
            model TEXT,
            input_tokens INTEGER DEFAULT 0,
            output_tokens INTEGER DEFAULT 0,
            total_tokens INTEGER DEFAULT 0,
            timestamp TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    """)
    conn.commit()

    # Seed default admin if no users exist
    cursor.execute("SELECT COUNT(*) FROM users")
    count = cursor.fetchone()[0]
    if count == 0:
        default_admin_email = "admin@mykloudz.com"
        default_admin_pass = "Admin@123"
        salt = secrets.token_hex(16)
        pw_hash = hash_password(default_admin_pass, salt)
        cursor.execute(
            "INSERT INTO users (email, password_hash, salt, role) VALUES (?, ?, ?, ?)",
            (default_admin_email, pw_hash, salt, "admin")
        )
        conn.commit()
    
    conn.close()


def verify_user(email: str, password: str) -> Optional[Dict[str, Any]]:
    """Verify user credentials and return user info dict if valid, else None."""
    email_clean = email.strip().lower()
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute(
        "SELECT email, password_hash, salt, role FROM users WHERE email = ?",
        (email_clean,)
    )
    row = cursor.fetchone()
    conn.close()
    
    if not row:
        return None
    
    db_email, stored_hash, salt, role = row
    if hash_password(password, salt) == stored_hash:
        return {
            "email": db_email,
            "role": role.lower()
        }
    return None


def add_user(email: str, password: str, role: str = "user") -> Tuple[bool, str]:
    """Add a new user with given email, password, and role."""
    email_clean = email.strip().lower()
    role_clean = role.strip().lower()
    if role_clean not in ["admin", "user"]:
        role_clean = "user"
        
    if not email_clean or "@" not in email_clean:
        return False, "Please provide a valid email address."
    if len(password) < 6:
        return False, "Password must be at least 6 characters long."
        
    salt = secrets.token_hex(16)
    pw_hash = hash_password(password, salt)
    
    try:
        conn = get_connection()
        cursor = conn.cursor()
        cursor.execute(
            "INSERT INTO users (email, password_hash, salt, role) VALUES (?, ?, ?, ?)",
            (email_clean, pw_hash, salt, role_clean)
        )
        conn.commit()
        conn.close()
        return True, f"User '{email_clean}' successfully created with role '{role_clean}'."
    except sqlite3.IntegrityError:
        return False, f"A user with email '{email_clean}' already exists."
    except Exception as e:
        return False, f"Database error: {str(e)}"


def update_user_password(email: str, new_password: str) -> Tuple[bool, str]:
    """Update password for an existing user."""
    email_clean = email.strip().lower()
    if len(new_password) < 6:
        return False, "Password must be at least 6 characters long."
        
    salt = secrets.token_hex(16)
    pw_hash = hash_password(new_password, salt)
    
    try:
        conn = get_connection()
        cursor = conn.cursor()
        cursor.execute(
            "UPDATE users SET password_hash = ?, salt = ? WHERE email = ?",
            (pw_hash, salt, email_clean)
        )
        if cursor.rowcount == 0:
            conn.close()
            return False, f"User '{email_clean}' not found."
        conn.commit()
        conn.close()
        return True, "Password updated successfully."
    except Exception as e:
        return False, f"Database error: {str(e)}"


def delete_user(email: str) -> Tuple[bool, str]:
    """Delete a user account."""
    email_clean = email.strip().lower()
    try:
        conn = get_connection()
        cursor = conn.cursor()
        # Protect against deleting the last admin
        cursor.execute("SELECT role FROM users WHERE email = ?", (email_clean,))
        row = cursor.fetchone()
        if not row:
            conn.close()
            return False, f"User '{email_clean}' not found."
            
        if row[0] == "admin":
            cursor.execute("SELECT COUNT(*) FROM users WHERE role = 'admin'")
            admin_count = cursor.fetchone()[0]
            if admin_count <= 1:
                conn.close()
                return False, "Cannot delete the only remaining admin account."
                
        cursor.execute("DELETE FROM users WHERE email = ?", (email_clean,))
        conn.commit()
        conn.close()
        return True, f"User '{email_clean}' deleted successfully."
    except Exception as e:
        return False, f"Database error: {str(e)}"


def get_all_users() -> List[Dict[str, Any]]:
    """Retrieve list of all users."""
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT id, email, role, created_at FROM users ORDER BY id ASC")
    rows = cursor.fetchall()
    conn.close()
    return [{"id": r[0], "email": r[1], "role": r[2], "created_at": r[3]} for r in rows]


def log_token_usage(user_email: str, file_name: str, model: str, input_tokens: int, output_tokens: int):
    """Log token consumption for a completed OCR document."""
    try:
        email_clean = user_email.strip().lower()
        inp = int(input_tokens or 0)
        out = int(output_tokens or 0)
        tot = inp + out
        
        conn = get_connection()
        cursor = conn.cursor()
        cursor.execute(
            "INSERT INTO token_usage (user_email, file_name, model, input_tokens, output_tokens, total_tokens) VALUES (?, ?, ?, ?, ?, ?)",
            (email_clean, file_name, model, inp, out, tot)
        )
        conn.commit()
        conn.close()
    except Exception as e:
        print(f"[auth_db] Failed to log token usage: {e}")


def get_token_usage_summary() -> pd.DataFrame:
    """Return aggregated token metrics grouped by user email."""
    conn = get_connection()
    query = """
        SELECT 
            user_email AS "User Email",
            COUNT(*) AS "Total Requests",
            SUM(input_tokens) AS "Input Tokens",
            SUM(output_tokens) AS "Output Tokens",
            SUM(total_tokens) AS "Total Tokens",
            MAX(timestamp) AS "Last Active"
        FROM token_usage
        GROUP BY user_email
        ORDER BY "Total Tokens" DESC
    """
    df = pd.read_sql_query(query, conn)
    conn.close()
    return df


def get_detailed_token_logs(limit: int = 500, filter_email: Optional[str] = None) -> pd.DataFrame:
    """Return recent granular token usage logs."""
    conn = get_connection()
    query = """
        SELECT 
            timestamp AS "Date & Time",
            user_email AS "User Email",
            file_name AS "File Name",
            model AS "Model",
            input_tokens AS "Input Tokens",
            output_tokens AS "Output Tokens",
            total_tokens AS "Total Tokens"
        FROM token_usage
    """
    params = []
    if filter_email and filter_email.strip():
        query += " WHERE user_email = ?"
        params.append(filter_email.strip().lower())
    query += " ORDER BY id DESC LIMIT ?"
    params.append(limit)
    
    df = pd.read_sql_query(query, conn, params=params)
    conn.close()
    return df


def get_user_token_stats(user_email: str) -> Dict[str, int]:
    """Get aggregated token consumption for a single user."""
    email_clean = user_email.strip().lower()
    try:
        conn = get_connection()
        cursor = conn.cursor()
        cursor.execute("""
            SELECT 
                COALESCE(SUM(total_tokens), 0),
                COALESCE(SUM(input_tokens), 0),
                COALESCE(SUM(output_tokens), 0),
                COUNT(*)
            FROM token_usage
            WHERE user_email = ?
        """, (email_clean,))
        row = cursor.fetchone()
        conn.close()
        return {
            "total_tokens": row[0] if row else 0,
            "input_tokens": row[1] if row else 0,
            "output_tokens": row[2] if row else 0,
            "total_requests": row[3] if row else 0,
        }
    except Exception:
        return {"total_tokens": 0, "input_tokens": 0, "output_tokens": 0, "total_requests": 0}


def get_global_token_stats() -> Dict[str, Any]:
    """Get organization-wide token and user stats."""
    try:
        conn = get_connection()
        cursor = conn.cursor()
        cursor.execute("""
            SELECT 
                COALESCE(SUM(total_tokens), 0),
                COALESCE(SUM(input_tokens), 0),
                COALESCE(SUM(output_tokens), 0),
                COUNT(*)
            FROM token_usage
        """)
        row = cursor.fetchone()
        
        cursor.execute("SELECT COUNT(*) FROM users")
        user_count = cursor.fetchone()[0]
        conn.close()
        
        return {
            "total_tokens": row[0] if row else 0,
            "input_tokens": row[1] if row else 0,
            "output_tokens": row[2] if row else 0,
            "total_requests": row[3] if row else 0,
            "total_users": user_count,
        }
    except Exception:
        return {"total_tokens": 0, "input_tokens": 0, "output_tokens": 0, "total_requests": 0, "total_users": 0}
