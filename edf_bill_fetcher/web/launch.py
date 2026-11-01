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
        "--init-user",
        help="Create a named account and exit; password is prompted or read from EDF_WEB_PASSWORD",
    )
    parser.add_argument("--source-root", action="append", default=[], metavar="NAME=PATH")
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
    store = Store(options.data)
    if options.init_user:
        store.user(
            options.init_user,
            os.environ.get("EDF_WEB_PASSWORD")
            or getpass.getpass("New password (12+ characters): "),
        )
        return
    if not store.list("user"):
        parser.error("Create an account first with edf-web --data PATH --init-user NAME")
    roots = {}
    for specification in options.source_root:
        name, separator, path = specification.partition("=")
        if not separator or not name or name in roots or not Path(path).is_dir():
            parser.error("Source roots must be unique NAME=PATH pairs pointing to existing folders")
        roots[name] = str(Path(path).resolve())
    uvicorn.run(
        create_app(options.data, roots=roots, secure_cookie=options.secure_cookie),
        host=options.host,
        port=options.port,
        workers=1,
        proxy_headers=False,
    )


if __name__ == "__main__":
    main()
