import { useEffect, useRef, useState } from "react";
import * as pdfjs from "pdfjs-dist";
import {
  EventBus,
  PDFFindController,
  PDFLinkService,
  PDFViewer,
} from "pdfjs-dist/web/pdf_viewer.mjs";
import "pdfjs-dist/web/pdf_viewer.css";

pdfjs.GlobalWorkerOptions.workerSrc = new URL(
  "pdfjs-dist/build/pdf.worker.min.mjs",
  import.meta.url,
).toString();

export default function Pdf({
  url,
  initialPage = 1,
}: {
  url: string;
  initialPage?: number;
}) {
  const container = useRef<HTMLDivElement>(null);
  const viewer = useRef<PDFViewer | null>(null);
  const bus = useRef<EventBus | null>(null);
  const [pages, setPages] = useState(0),
    [page, setPage] = useState(1),
    [error, setError] = useState(""),
    [query, setQuery] = useState("");
  useEffect(() => {
    setError("");
    setPages(0);
    const element = container.current!;
    const eventBus = new EventBus();
    bus.current = eventBus;
    const linkService = new PDFLinkService({ eventBus });
    const findController = new PDFFindController({ eventBus, linkService });
    const display = new PDFViewer({
      container: element,
      viewer: element.firstElementChild as HTMLDivElement,
      eventBus,
      linkService,
      findController,
    });
    viewer.current = display;
    linkService.setViewer(display);
    const task = pdfjs.getDocument({
      url,
      withCredentials: true,
      cMapUrl: "/pdf/cmaps/",
      cMapPacked: true,
      standardFontDataUrl: "/pdf/standard_fonts/",
      wasmUrl: "/pdf/wasm/",
    });
    let active = true;
    eventBus.on("pagesinit", () => {
      display.currentScaleValue = "page-width";
      display.currentPageNumber = Math.min(initialPage, display.pagesCount);
      setPages(display.pagesCount);
    });
    eventBus.on("pagechanging", (event: { pageNumber: number }) =>
      setPage(event.pageNumber),
    );
    task.promise
      .then((document) => {
        if (active) {
          display.setDocument(document);
          linkService.setDocument(document);
        }
      })
      .catch((error) => {
        if (active) setError(String(error));
      });
    return () => {
      active = false;
      display.cleanup();
      void task.destroy();
    };
  }, [url, initialPage]);
  return (
    <section className="pdf-shell" aria-label="Original PDF">
      <div className="pdf-toolbar">
        <button
          disabled={page <= 1}
          onClick={() => {
            if (viewer.current) viewer.current.currentPageNumber--;
          }}
        >
          Previous
        </button>
        <label>
          Page{" "}
          <input
            aria-label="PDF page"
            type="number"
            min={1}
            max={pages || 1}
            value={page}
            onChange={(event) => {
              if (viewer.current && pages)
                viewer.current.currentPageNumber = Math.max(
                  1,
                  Math.min(pages, Number(event.target.value)),
                );
            }}
          />
        </label>
        <span>of {pages}</span>
        <button
          disabled={page >= pages}
          onClick={() => {
            if (viewer.current) viewer.current.currentPageNumber++;
          }}
        >
          Next
        </button>
        <select
          aria-label="PDF zoom"
          onChange={(event) => {
            if (viewer.current)
              viewer.current.currentScaleValue = event.target.value;
          }}
        >
          <option value="page-width">Fit width</option>
          <option value="page-fit">Fit page</option>
          <option value="1">100%</option>
          <option value="1.5">150%</option>
          <option value="2">200%</option>
        </select>
        <input
          aria-label="Find in original"
          placeholder="Find in original…"
          value={query}
          onChange={(event) => setQuery(event.target.value)}
        />
        <button
          onClick={() =>
            bus.current?.dispatch("find", {
              source: viewer.current,
              type: "again",
              query,
              phraseSearch: true,
              caseSensitive: false,
              entireWord: false,
              highlightAll: true,
              findPrevious: false,
            })
          }
        >
          Find next
        </button>
      </div>
      {error && <p role="alert">{error}</p>}
      <div ref={container} className="pdf-container">
        <div className="pdfViewer" />
      </div>
    </section>
  );
}
