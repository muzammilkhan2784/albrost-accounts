"""Add a login or change its password. The password is only stored as a hash.

Usage:  python set_password.py USERNAME ROLE "Display Name"
        ROLE is owner (full access) or manager (can only add new entries).
"""
import getpass
import json
import sys

from server import DATA, USERS_FILE, hash_password


def main():
    if len(sys.argv) != 4 or sys.argv[2] not in ("owner", "manager"):
        raise SystemExit(__doc__)
    username, role, name = sys.argv[1:]
    prompt = f"Password for {username}: "
    password = getpass.getpass(prompt) if sys.stdin.isatty() else sys.stdin.readline().rstrip("\r\n")
    if not password:
        raise SystemExit("Password can't be empty.")
    DATA.mkdir(parents=True, exist_ok=True)
    users = json.loads(USERS_FILE.read_text(encoding="utf-8")) if USERS_FILE.exists() else {}
    salt, digest = hash_password(password)
    users[username] = {"name": name, "role": role, "salt": salt, "hash": digest}
    USERS_FILE.write_text(json.dumps(users, indent=2), encoding="utf-8")
    print(f"Saved login for {username} ({role}).")


if __name__ == "__main__":
    main()
