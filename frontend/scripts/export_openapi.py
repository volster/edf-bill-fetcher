"""Generate a path-independent HTTP schema without starting a server or worker."""

import json
from pathlib import Path
from tempfile import TemporaryDirectory

from edf_bill_fetcher.web.api import create_app

with TemporaryDirectory(prefix="edf-schema-") as directory:
    app = create_app(
        directory,
        start_worker=False,
        database_url="sqlite:///" + str(Path(directory) / "catalogue.sqlite"),
    )
    Path("openapi.json").write_text(json.dumps(app.openapi(), indent=2) + "\n", encoding="utf-8")
    app.state.store.engine.dispose()
