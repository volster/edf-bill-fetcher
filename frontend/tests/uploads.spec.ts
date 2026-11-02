import { createHash } from "node:crypto";
import { expect, test, type Page } from "@playwright/test";

async function build(page: Page, name: string) {
  await page.goto("/");
  await page.getByLabel("Username", { exact: true }).fill("reviewer");
  await page
    .getByLabel("Password", { exact: true })
    .fill("synthetic-browser-password");
  await page.getByRole("button", { name: "Sign in", exact: true }).click();
  await page.getByRole("link", { name: "Build a pack", exact: true }).click();
  await page.getByLabel("Source label", { exact: true }).fill(name);
}

for (const lostResponse of ["chunk", "seal"]) {
  test(`resume a multi-chunk source after losing its ${lostResponse} response`, async ({
    page,
  }) => {
    const name = `Upload recovery ${lostResponse}`;
    await build(page, name);
    const bytes = Buffer.alloc(9 * 1024 * 1024 + 17, 73);
    const files = [
      {
        name: "archive.pst",
        mimeType: "application/octet-stream",
        buffer: bytes,
      },
      { name: "empty.txt", mimeType: "text/plain", buffer: Buffer.alloc(0) },
    ];
    let interrupted = false;
    const chunks: number[] = [];
    await page.route("**/api/assets/*/content?**", async (route) => {
      const offset = Number(
        new URL(route.request().url()).searchParams.get("offset"),
      );
      chunks.push(offset);
      if (
        lostResponse === "chunk" &&
        offset === 4 * 1024 * 1024 &&
        !interrupted
      ) {
        // WebKit interception may omit a Blob request's body. Send the exact
        // selected bytes so this models a lost acknowledgement after a real write.
        const response = await route.fetch({
          postData: bytes.subarray(offset, offset + 4 * 1024 * 1024),
        });
        expect(response.status()).toBe(200);
        expect(await response.json()).toEqual({
          offset: offset + 4 * 1024 * 1024,
        });
        interrupted = true;
        await route.abort("failed");
      } else await route.continue();
    });
    if (lostResponse === "seal")
      await page.route("**/api/assets/*/seal", async (route) => {
        if (!interrupted) {
          const response = await route.fetch();
          expect(response.status()).toBe(200);
          interrupted = true;
          await route.abort("failed");
        } else await route.continue();
      });
    const input = page.getByLabel("Choose files or ZIPs", { exact: true });
    await input.setInputFiles(files);
    await page
      .getByRole("button", { name: "Upload and add", exact: true })
      .click();
    await expect(page.getByRole("alert")).toBeVisible();
    expect(interrupted).toBe(true);
    const chunksBefore = chunks.length;
    if (lostResponse === "chunk") {
      await input.setInputFiles([files[0]]);
      await page
        .getByRole("button", { name: "Resume pending upload", exact: true })
        .click();
      await expect(page.getByRole("alert")).toContainText(
        "same paths and sizes",
      );
      expect(chunks).toHaveLength(chunksBefore);
      await input.setInputFiles(files);
    }
    await page
      .getByRole("button", { name: "Resume pending upload", exact: true })
      .click();
    await expect(page.getByText("Source ready", { exact: true })).toBeVisible();
    expect(chunks.slice(chunksBefore)).toEqual(
      lostResponse === "chunk" ? [8 * 1024 * 1024, 0] : [],
    );
    const assets = await (await page.request.get("/api/assets")).json();
    const uploaded = assets.filter(
      (asset: { data: { name: string } }) => asset.data.name === name,
    );
    expect(uploaded).toHaveLength(1);
    expect(uploaded[0].data.sealed).toBe(true);
    for (const file of files) {
      const response = await page.request.get(
        `/api/assets/${uploaded[0].id}/offset?path=${file.name}`,
      );
      expect(await response.json()).toEqual({
        offset: file.buffer.length,
        sha256: createHash("sha256").update(file.buffer).digest("hex"),
      });
    }
    // A completed input must reset so selecting the same originals emits change again.
    await expect(input).toHaveValue("");
    await input.setInputFiles(files);
    await expect(
      page.getByRole("button", { name: "Upload and add", exact: true }),
    ).toBeEnabled();
  });
}

test("pause an upload and resume the same source without creating another asset", async ({
  page,
}) => {
  const name = "Paused browser upload";
  await build(page, name);
  let started!: () => void;
  const requestStarted = new Promise<void>((resolve) => {
    started = resolve;
  });
  let release!: () => void;
  const resumeRequest = new Promise<void>((resolve) => {
    release = resolve;
  });
  await page.route("**/api/assets/*/content?**", async (route) => {
    started();
    await resumeRequest;
    await route.abort("failed");
  });
  await page.getByLabel("Choose files or ZIPs", { exact: true }).setInputFiles({
    name: "paused.pdf",
    mimeType: "application/pdf",
    buffer: Buffer.from("synthetic paused source"),
  });
  await page
    .getByRole("button", { name: "Upload and add", exact: true })
    .click();
  await requestStarted;
  await expect(page.getByLabel("Source type", { exact: true })).toBeDisabled();
  await page.getByRole("button", { name: "Pause upload", exact: true }).click();
  await expect(page.getByRole("alert")).toContainText("Upload paused");
  release();
  await page.unrouteAll({ behavior: "wait" });
  await page
    .getByRole("button", { name: "Resume pending upload", exact: true })
    .click();
  await expect(page.getByText("Source ready", { exact: true })).toBeVisible();
  const assets = await (await page.request.get("/api/assets")).json();
  expect(
    assets.filter(
      (asset: { data: { name: string } }) => asset.data.name === name,
    ),
  ).toHaveLength(1);
});
