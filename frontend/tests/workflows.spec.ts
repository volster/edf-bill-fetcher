import { execFileSync } from "node:child_process";
import { mkdtempSync, rmSync } from "node:fs";
import { tmpdir } from "node:os";
import { join, resolve } from "node:path";
import { test, expect, type Page } from "@playwright/test";

async function signIn(page: Page) {
  await page.goto("/");
  await page.getByLabel("Username", { exact: true }).fill("reviewer");
  await page
    .getByLabel("Password", { exact: true })
    .fill("synthetic-browser-password");
  await page.getByRole("button", { name: "Sign in", exact: true }).click();
  await expect(
    page.getByRole("heading", { name: "Runs & evidence" }),
  ).toBeVisible();
}

test("review a saved analysis, reject stale edits, undo and preserve pending forms", async ({
  page,
}) => {
  const errors: string[] = [];
  page.on("pageerror", (error) => errors.push(error.message));
  await signIn(page);
  await page
    .getByRole("heading", { name: "Synthetic baseline", exact: true })
    .click();
  await page.getByRole("button", { name: "BROWSER-001", exact: true }).click();
  await page.getByLabel("Period Charge (£)", { exact: true }).fill("80.00");
  await page
    .getByLabel("Reason for change", { exact: true })
    .fill("Verified against synthetic original");
  await page
    .getByRole("button", { name: "Apply to draft", exact: true })
    .click();
  await expect(
    page.getByText("Draft revision 1", { exact: false }),
  ).toBeVisible();
  await expect(
    page.getByLabel("Period Charge (£)", { exact: true }),
  ).toHaveValue("80");
  await page.getByRole("button", { name: "Undo", exact: true }).click();
  await expect(
    page.getByLabel("Period Charge (£)", { exact: true }),
  ).toHaveValue("100");
  await page.getByRole("button", { name: "Redo", exact: true }).click();
  await expect(
    page.getByLabel("Period Charge (£)", { exact: true }),
  ).toHaveValue("80");
  await page.getByLabel("Period Charge (£)", { exact: true }).fill("70");
  page.once("dialog", (dialog) => dialog.dismiss());
  await page
    .getByRole("tab", { name: "Claims & complaint", exact: true })
    .click();
  await expect(
    page.getByLabel("Period Charge (£)", { exact: true }),
  ).toHaveValue("70");
  await page.getByRole("button", { name: "Discard form edits" }).click();
  await page
    .getByRole("tab", { name: "Claims & complaint", exact: true })
    .click();
  await page
    .getByLabel("Complaint summary", { exact: true })
    .fill("Synthetic disputed rebill");
  await page
    .getByLabel("Requested outcome", { exact: true })
    .fill("Review the unbilled increase");
  await page.getByRole("button", { name: "Save complaint wording" }).click();
  await expect(
    page.getByText("Draft revision 4", { exact: false }),
  ).toBeVisible();
  await page.reload();
  await page
    .getByRole("tab", { name: "Claims & complaint", exact: true })
    .click();
  await expect(
    page.getByLabel("Complaint summary", { exact: true }),
  ).toHaveValue("Synthetic disputed rebill");
  await page.getByRole("tab", { name: "Dispute ledger", exact: true }).click();
  await page
    .getByLabel("Bill", { exact: true })
    .selectOption({ label: "BROWSER-001 · 01/02/2024 · A-123" });
  await expect(
    page.getByText("Unapplied form edits", { exact: true }),
  ).not.toBeVisible();
  await page.getByLabel("Status", { exact: true }).selectOption("challenged");
  await page.getByLabel("Challenged Amount", { exact: true }).fill("80.00");
  await page.getByLabel("Proposed Reduction", { exact: true }).fill("40.00");
  await page
    .getByLabel("Basis", { exact: true })
    .fill("Synthetic review position");
  await page
    .getByRole("button", { name: "Save position", exact: true })
    .click();
  await expect(
    page.getByText("Draft revision 5", { exact: false }),
  ).toBeVisible();
  await page
    .getByRole("button", { name: "Clear position", exact: true })
    .click();
  await expect(
    page.getByText("Draft revision 6", { exact: false }),
  ).toBeVisible();
  await expect(
    page.getByLabel("Challenged Amount", { exact: true }),
  ).toHaveValue("");
  await page.getByRole("button", { name: "Undo", exact: true }).click();
  await expect(
    page.getByLabel("Challenged Amount", { exact: true }),
  ).toHaveValue("80.00");
  await page.getByRole("button", { name: "Redo", exact: true }).click();
  await expect(
    page.getByLabel("Challenged Amount", { exact: true }),
  ).toHaveValue("");
  await expect(
    page.getByText("Unapplied form edits", { exact: true }),
  ).not.toBeVisible();
  expect(errors).toEqual([]);
});

test("extract from a configured server folder, inspect original PDF and download a portable pack", async ({
  page,
}) => {
  const errors: string[] = [];
  page.on("pageerror", (error) => errors.push(error.message));
  await signIn(page);
  const navigationAppearance = () =>
    page
      .getByRole("navigation", { name: "Main navigation" })
      .evaluate((node) => {
        const appearance = getComputedStyle(node.closest("aside")!);
        return {
          background: appearance.backgroundColor,
          color: appearance.color,
          padding: appearance.padding,
          borderRadius: appearance.borderRadius,
        };
      });
  const appearanceBeforePdf = await navigationAppearance();

  await page.getByRole("link", { name: "Build a pack", exact: true }).click();
  await page
    .getByLabel("Pack name", { exact: true })
    .fill("Browser extraction");
  await page.getByLabel("Location", { exact: true }).selectOption("server");
  await page
    .getByLabel("Server source folder", { exact: true })
    .selectOption("Synthetic bills");
  // Folders select paths below a configured root; a single PDF can be selected
  // by its containing folder on upload. Use an explicit HTM-free nested folder.
  await page
    .getByLabel("Relative folder or filename", { exact: true })
    .fill("./");
  await page
    .getByRole("button", { name: "Add selected path", exact: true })
    .click();
  await page
    .getByRole("button", { name: "Inspect everything", exact: true })
    .click();
  await expect(
    page.getByRole("heading", { name: "Intake inventory", exact: true }),
  ).toBeVisible();
  await page
    .getByRole("button", { name: "Create evidence pack", exact: true })
    .click();
  await expect(
    page.getByRole("heading", { name: "Browser extraction", exact: true }),
  ).toBeVisible();
  await expect(
    page.getByRole("button", { name: "KI-12345", exact: true }),
  ).toBeVisible({ timeout: 60_000 });
  const runId = page.url().split("/runs/")[1];
  let finishing = false;
  let unblockArtifacts!: () => void;
  let artifactsRequested!: () => void;
  const artifactGate = new Promise<void>((resolve) => {
    unblockArtifacts = resolve;
  });
  const requested = new Promise<void>((resolve) => {
    artifactsRequested = resolve;
  });
  await page.route(`**/api/jobs/${runId}`, async (route) => {
    const response = await route.fetch();
    const job = await response.json();
    await route.fulfill({
      response,
      json: {
        ...job,
        data: { ...job.data, status: finishing ? job.data.status : "running" },
      },
    });
  });
  await page.route(`**/api/jobs/${runId}/artifacts`, async (route) => {
    const response = await route.fetch();
    if (finishing) {
      artifactsRequested();
      await artifactGate;
    }
    await route.fulfill({ response });
  });
  await page.reload();
  await page.getByRole("button", { name: "KI-12345", exact: true }).click();
  await expect(page.getByLabel("PDF page", { exact: true })).toBeVisible();
  finishing = true;
  await requested;
  try {
    // A final artifact refresh must retain the selected record and its PDF.
    await expect(page.getByLabel("PDF page", { exact: true })).toBeVisible();
  } finally {
    unblockArtifacts();
  }
  await expect(page.locator(".pdfViewer .page canvas").first()).toBeVisible();
  await expect.poll(navigationAppearance).toEqual(appearanceBeforePdf);

  await page.getByRole("tab", { name: "Outputs", exact: true }).click();
  const download = page.waitForEvent("download");
  await page
    .getByRole("link", { name: "Download complete evidence pack", exact: true })
    .click();
  expect((await download).suggestedFilename()).toBe("Evidence_pack.zip");
  expect(errors).toEqual([]);
});

test("small screens retain navigation and inline sign-in errors", async ({
  page,
}) => {
  await page.setViewportSize({ width: 390, height: 844 });
  await page.goto("/");
  await page.getByLabel("Username", { exact: true }).fill("reviewer");
  await page.getByLabel("Password", { exact: true }).fill("wrong");
  await page.getByRole("button", { name: "Sign in", exact: true }).click();
  await expect(page.getByRole("alert")).toContainText(
    "Invalid username or password",
  );
  await page
    .getByLabel("Password", { exact: true })
    .fill("synthetic-browser-password");
  await page.getByRole("button", { name: "Sign in", exact: true }).click();
  await expect(
    page.getByRole("navigation", { name: "Main navigation" }),
  ).toBeVisible();
});

test("mixed folder uploads include ZIP members and retain unsupported-file decisions", async ({
  page,
}) => {
  const temporary = mkdtempSync(join(tmpdir(), "edf-upload-browser-"));
  try {
    execFileSync(
      process.env.EDF_TEST_PYTHON || "python",
      [
        "-c",
        "import sys; from pathlib import Path; from tests.test_mixed_intake import mixed_folder; mixed_folder(Path(sys.argv[1]))",
        join(temporary, "mixed"),
      ],
      { cwd: resolve("..") },
    );
    await signIn(page);
    await page.getByRole("link", { name: "Build a pack", exact: true }).click();
    await page
      .getByLabel("Pack name", { exact: true })
      .fill("Mixed browser upload");
    await page
      .getByLabel("Upload selection", { exact: true })
      .selectOption("folder");
    await page
      .getByLabel("Choose a folder", { exact: true })
      .setInputFiles(join(temporary, "mixed"));
    await page
      .getByRole("button", { name: "Upload and add", exact: true })
      .click();
    await expect(page.getByText("Source ready", { exact: true })).toBeVisible();
    await page
      .getByRole("button", { name: "Inspect everything", exact: true })
      .click();
    const inventory = page.getByRole("table", {
      name: "Intake files",
      exact: true,
    });
    await expect(inventory).toContainText("more.zip!/2024/second.pdf");
    await expect(inventory).toContainText("photo.png");
    await expect(
      page.getByRole("button", { name: "Create evidence pack", exact: true }),
    ).toBeDisabled();
    await page
      .getByLabel("I have checked the unprocessed files", { exact: false })
      .check();
    // Select only Excel for this test; the other browser scenario exercises all formats.
    await page.getByText("Report formats & sections", { exact: true }).click();
    for (const format of ["PDF", "DOCX", "HTML"])
      await page.getByLabel(format, { exact: true }).uncheck();
    await page
      .getByRole("button", { name: "Create evidence pack", exact: true })
      .click();
    await expect(
      page.getByRole("heading", { name: "Mixed browser upload", exact: true }),
    ).toBeVisible();
    await expect(
      page.getByRole("button", { name: "KI-10001", exact: true }),
    ).toBeVisible({ timeout: 60_000 });
    await expect(
      page.getByRole("button", { name: "KI-10002", exact: true }),
    ).toBeVisible();
    await page.getByRole("tab", { name: "Outputs", exact: true }).click();
    await expect(
      page.getByRole("link", { name: "intake_inventory.json", exact: true }),
    ).toBeVisible();
  } finally {
    rmSync(temporary, { recursive: true, force: true });
  }
});

test("expired sessions keep unapplied record edits until reauthentication", async ({
  page,
}) => {
  await signIn(page);
  await page
    .getByRole("heading", { name: "Synthetic baseline", exact: true })
    .click();
  await page.getByRole("button", { name: "BROWSER-001", exact: true }).click();
  await page.getByLabel("Period Charge (£)", { exact: true }).fill("66.25");
  await page
    .getByLabel("Reason for change", { exact: true })
    .fill("Preserved during expiry");
  await page.context().clearCookies();
  await page
    .getByRole("button", { name: "Apply to draft", exact: true })
    .click();
  const overlay = page.getByRole("dialog", {
    name: "Session expired",
    exact: true,
  });
  await expect(overlay).toBeVisible();
  await overlay.getByLabel("Username", { exact: true }).fill("reviewer");
  await overlay
    .getByLabel("Password", { exact: true })
    .fill("synthetic-browser-password");
  await overlay.getByRole("button", { name: "Sign in", exact: true }).click();
  await expect(overlay).not.toBeVisible();
  await expect(
    page.getByLabel("Period Charge (£)", { exact: true }),
  ).toHaveValue("66.25");
  await page
    .getByRole("button", { name: "Apply to draft", exact: true })
    .click();
  await page.reload();
  await page.getByRole("button", { name: "BROWSER-001", exact: true }).click();
  await expect(
    page.getByLabel("Period Charge (£)", { exact: true }),
  ).toHaveValue("66.25");
});

test("interrupted uploads reject a changed original and resume from the verified chunk", async ({
  page,
}) => {
  await signIn(page);
  await page.getByRole("link", { name: "Build a pack", exact: true }).click();
  const buffer = Buffer.alloc(4 * 1024 * 1024 + 512, 120);
  const file = {
    name: "large-original.bin",
    mimeType: "application/octet-stream",
    buffer,
  };
  const selection = page.getByLabel("Choose files or ZIPs", { exact: true });
  await selection.setInputFiles(file);
  const offsets: number[] = [];
  let interrupted = false;
  await page.route("**/api/assets/*/content?*", async (route) => {
    const offset = Number(
      new URL(route.request().url()).searchParams.get("offset"),
    );
    offsets.push(offset);
    if (offset > 0 && !interrupted) {
      interrupted = true;
      await route.abort("failed");
    } else await route.continue();
  });
  await page
    .getByRole("button", { name: "Upload and add", exact: true })
    .click();
  const resume = page.getByRole("button", {
    name: "Resume pending upload",
    exact: true,
  });
  await expect(resume).toBeEnabled();
  const changed = Buffer.from(buffer);
  changed[0] = 121;
  await selection.setInputFiles({ ...file, buffer: changed });
  await resume.click();
  await expect(page.getByRole("alert")).toContainText(
    "does not match the pending upload",
  );
  await selection.setInputFiles(file);
  await resume.click();
  await expect(page.getByText("Source ready", { exact: true })).toBeVisible();
  expect(offsets).toEqual([0, 4 * 1024 * 1024, 4 * 1024 * 1024]);
  expect(
    await page.evaluate(() => sessionStorage.getItem("edf-upload")),
  ).toBeNull();
});
