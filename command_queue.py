"""Persistent command admission and at-most-once Outlook submission."""
import sqlite3
import threading
import time
from contextlib import closing


class CommandQueue:
    def __init__(self, path):
        self.path = path
        self.lock = threading.RLock()
        with closing(sqlite3.connect(path)) as db, db:
            db.executescript("""
                CREATE TABLE IF NOT EXISTS mail_switch (
                    singleton INTEGER PRIMARY KEY CHECK(singleton=1),
                    active INTEGER NOT NULL, since REAL NOT NULL
                );
                CREATE TABLE IF NOT EXISTS mail_command (
                    chat TEXT, message TEXT, state TEXT NOT NULL,
                    PRIMARY KEY(chat, message)
                );
            """)
            db.execute("INSERT OR IGNORE INTO mail_switch VALUES (1,1,?)", (time.time(),))
            db.execute("UPDATE mail_switch SET since=? WHERE singleton=1", (time.time(),))
            db.execute("UPDATE mail_command SET state='cancelled' WHERE state='pending'")

    def switch(self, active, since):
        with self.lock, closing(sqlite3.connect(self.path)) as db, db:
            if not active:
                db.execute("UPDATE mail_command SET state='cancelled' WHERE state='pending'")
            db.execute("UPDATE mail_switch SET active=?, since=? WHERE active!=?",
                       (active, since, active))

    def admit(self, chat, message, timestamp):
        if not chat or not message or timestamp is None:
            return False
        with self.lock, closing(sqlite3.connect(self.path)) as db, db:
            return db.execute("""
                INSERT OR IGNORE INTO mail_command
                SELECT ?, ?, 'pending' FROM mail_switch WHERE active=1 AND since<=?
            """, (chat, message, timestamp)).rowcount == 1

    def cancel(self, chat, message):
        with self.lock, closing(sqlite3.connect(self.path)) as db, db:
            db.execute("UPDATE mail_command SET state='cancelled' WHERE chat=? AND message=? AND state='pending'",
                       (chat, message))

    def submit(self, chat, message, deliver):
        with self.lock, closing(sqlite3.connect(self.path)) as db:
            with db:
                claimed = db.execute("""
                    UPDATE mail_command SET state='unconfirmed'
                    WHERE chat=? AND message=? AND state='pending'
                    AND EXISTS (SELECT 1 FROM mail_switch WHERE active=1)
                """, (chat, message)).rowcount
            if not claimed:
                return False
            # Commit before Outlook is called. An ambiguous outcome requires review.
            if deliver():
                with db:
                    db.execute("UPDATE mail_command SET state='submitted' WHERE chat=? AND message=?",
                               (chat, message))
                return True
            return False
