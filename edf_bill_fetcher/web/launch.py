"""Launch the same authenticated web application from source or a container."""

from __future__ import annotations

import argparse
import getpass
import os
from pathlib import Path


def main(argv: list[str] | None = None) -> None:
    """Capture server settings before importing optional HTTP dependencies."""
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--data", default=os.environ.get("EDF_WEB_DATA", "web-data"))
    parser.add_argument("--host", default="127.0.0.1")
    parser.add_argument("--port", type=int, default=8000)
    parser.add_argument(
        "--database-url", help="PostgreSQL/SQLite URL; defaults to EDF_DATABASE_URL or local SQLite"
    )
    parser.add_argument(
        "--init-user",
        help="Create a named account and exit; password is prompted or read from EDF_WEB_PASSWORD",
    )
    parser.add_argument("--source-root", action="append", default=[], metavar="NAME=PATH")
    parser.add_argument(
        "--reset-password",
        metavar="NAME",
        help="Reset an existing account's password and revoke its sessions",
    )
    parser.add_argument(
        "--secure-cookie",
        action="store_true",
        help="Require HTTPS for session cookies behind a local TLS proxy",
    )
    options = parser.parse_args(argv)
    try:
        import uvicorn

        from edf_bill_fetcher.web.api import create_app
        from edf_bill_fetcher.web.storage import Store
    except ImportError as error:
        parser.error(f"Install the web extra first: pip install '.[web,msg]' ({error})")
    store = Store(options.data, options.database_url)
    store.import_sqlite()
    if options.reset_password:
        from edf_bill_fetcher.web.accounts import username

        user = next(
            (
                user
                for user in store.accounts.users()
                if user["username"] == username(options.reset_password)
            ),
            None,
        )
        if user is None:
            parser.error("Account not found")
        store.accounts.credentials(
            user["id"],
            user["username"],
            os.environ.get("EDF_WEB_PASSWORD") or getpass.getpass("New password: "),
            reset=True,
        )
        store.engine.dispose()
        return
    if options.init_user:
        store.user(
            options.init_user,
            os.environ.get("EDF_WEB_PASSWORD") or getpass.getpass("New password: "),
        )
        store.engine.dispose()
        return
    store.engine.dispose()
    roots = {}
    for specification in options.source_root:
        name, separator, path = specification.partition("=")
        if not separator or not name or name in roots or not Path(path).is_dir():
            parser.error("Source roots must be unique NAME=PATH pairs pointing to existing folders")
        roots[name] = str(Path(path).resolve())
    try:
        application = create_app(
            options.data,
            roots=roots,
            secure_cookie=options.secure_cookie,
            database_url=options.database_url,
        )
    except ValueError as error:
        parser.error(str(error))
    uvicorn.run(
        application,
        host=options.host,
        port=options.port,
        workers=1,
        proxy_headers=False,
    )


if __name__ == "__main__":
    main()
