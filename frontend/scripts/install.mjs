import { cpSync, mkdirSync, rmSync } from "node:fs";
const target = new URL("../../edf_bill_fetcher/web/static/", import.meta.url);
rmSync(target, { recursive: true, force: true });
mkdirSync(target, { recursive: true });
cpSync(new URL("../dist/", import.meta.url), target, { recursive: true });

for (const folder of ["cmaps", "standard_fonts", "wasm"]) {
  cpSync(
    new URL("../node_modules/pdfjs-dist/" + folder + "/", import.meta.url),
    new URL("pdf/" + folder + "/", target),
    { recursive: true },
  );
}
