import { test, expect, type Page } from "@playwright/test";

async function signIn(page: Page, username: string, password: string) {
  await page.goto("/");
  await page.getByLabel("Username", { exact: true }).fill(username);
  await page.getByLabel("Password", { exact: true }).fill(password);
  await page.getByRole("button", { name: "Sign in", exact: true }).click();
}

test("default credentials lead to account setup and accept a user-chosen short password", async ({
  page,
}) => {
  await signIn(page, "admin", "admin");
  await expect(
    page.getByRole("heading", { name: "Set up your account" }),
  ).toBeVisible();
  await expect(
    page.getByRole("navigation", { name: "Main navigation" }),
  ).toHaveCount(0);
  await page.getByLabel("Username", { exact: true }).fill("browser-owner");
  await page.getByLabel("New password", { exact: true }).fill("a");
  await page.getByLabel("Confirm password", { exact: true }).fill("b");
  await page.getByRole("button", { name: "Finish setup", exact: true }).click();
  await expect(page.getByRole("alert")).toContainText("passwords do not match");
  await page.getByLabel("Confirm password", { exact: true }).fill("a");
  await page.getByRole("button", { name: "Finish setup", exact: true }).click();
  await expect(
    page.getByRole("heading", { name: "Runs & evidence", exact: true }),
  ).toBeVisible();
  page.once("dialog", (dialog) => dialog.accept());
  await page.getByRole("button", { name: "Sign out", exact: true }).click();
  await signIn(page, "admin", "admin");
  await expect(page.getByRole("alert")).toContainText(
    "Invalid username or password",
  );
  await signIn(page, "browser-owner", "a");
  await expect(
    page.getByRole("heading", { name: "Runs & evidence", exact: true }),
  ).toBeVisible();
});

test("share a workspace, enforce viewer controls and remove access from an active browser", async ({
  page,
  browser,
}) => {
  await signIn(page, "reviewer", "synthetic-browser-password");
  await expect(
    page.getByRole("heading", { name: "Runs & evidence", exact: true }),
  ).toBeVisible();
  const workspaces = await (await page.request.get("/api/workspaces")).json();
  const shared = workspaces[0].id as string;
  await page
    .getByRole("link", { name: "Workspaces & sharing", exact: true })
    .click();
  await page.getByLabel("Member username", { exact: true }).fill("viewer");
  await page
    .getByLabel("Member access", { exact: true })
    .selectOption("viewer");
  await page.getByRole("button", { name: "Add member", exact: true }).click();
  await expect(
    page.getByLabel("Access for viewer", { exact: true }),
  ).toHaveValue("viewer");
  const context = await browser.newContext();
  const reader = await context.newPage();
  try {
    await signIn(reader, "viewer", "v");
    await expect(
      reader.getByRole("heading", { name: "Runs & evidence", exact: true }),
    ).toBeVisible();
    await expect(
      reader.getByRole("heading", { name: "Synthetic baseline", exact: true }),
    ).toHaveCount(0);
    reader.once("dialog", (dialog) => dialog.accept());
    await reader
      .getByLabel("Active workspace", { exact: true })
      .selectOption(shared);
    await expect(
      reader.getByText("Viewer access", { exact: false }),
    ).toBeVisible();
    await expect(
      reader.getByRole("link", { name: "Build a pack", exact: true }),
    ).toHaveCount(0);
    await reader
      .getByRole("heading", { name: "Synthetic baseline", exact: true })
      .click();
    await reader
      .getByRole("button", { name: "BROWSER-001", exact: true })
      .click();
    await expect(
      reader.getByRole("button", { name: "Apply to draft", exact: true }),
    ).toBeDisabled();
    await expect(
      reader.getByLabel("Period Charge (£)", { exact: true }),
    ).toBeDisabled();
    await page
      .getByLabel("Access for viewer", { exact: true })
      .selectOption("editor");
    await expect(
      reader.getByRole("button", { name: "Apply to draft", exact: true }),
    ).toBeEnabled();
    const row = page
      .locator(".member-row")
      .filter({ has: page.getByLabel("Access for viewer", { exact: true }) });
    page.once("dialog", (dialog) => dialog.accept());
    await row
      .getByRole("button", { name: "Remove member", exact: true })
      .click();
    await expect(
      reader.getByRole("heading", { name: "Runs & evidence", exact: true }),
    ).toBeVisible();
    await expect(
      reader.getByRole("heading", { name: "Synthetic baseline", exact: true }),
    ).toHaveCount(0);
    await expect(
      reader
        .getByLabel("Active workspace", { exact: true })
        .locator("option", { hasText: "reviewer workspace" }),
    ).toHaveCount(0);
  } finally {
    await context.close();
  }
});

test("an older sharing form cannot overwrite changes after the workspace list refreshes", async ({
  page,
  context,
}) => {
  await signIn(page, "reviewer", "synthetic-browser-password");
  await expect(
    page.getByRole("heading", { name: "Runs & evidence", exact: true }),
  ).toBeVisible();
  await page
    .getByRole("link", { name: "Workspaces & sharing", exact: true })
    .click();
  const stale = await context.newPage();
  await stale.goto("/#/sharing");
  await expect(
    stale.getByLabel("Member username", { exact: true }),
  ).toBeVisible();
  const refreshed = stale.waitForResponse(
    (response) =>
      response.url().endsWith("/api/workspaces") &&
      response.request().resourceType() === "fetch",
  );
  await page.getByLabel("Member username", { exact: true }).fill("viewer");
  await page.getByRole("button", { name: "Add member", exact: true }).click();
  await expect(
    page.getByLabel("Access for viewer", { exact: true }),
  ).toBeVisible();
  await refreshed;
  await stale
    .getByLabel("Member username", { exact: true })
    .fill("browser-owner");
  await stale.getByRole("button", { name: "Add member", exact: true }).click();
  await expect(stale.getByRole("alert")).toContainText(
    "Workspace sharing changed",
  );
  stale.once("dialog", (dialog) => dialog.accept());
  await stale
    .getByRole("button", { name: "Reload sharing", exact: true })
    .click();
  await expect(
    stale.getByLabel("Member username", { exact: true }),
  ).toHaveValue("");
  await stale
    .getByLabel("Member username", { exact: true })
    .fill("browser-owner");
  await stale.getByRole("button", { name: "Add member", exact: true }).click();
  await expect(
    stale.getByLabel("Access for browser-owner", { exact: true }),
  ).toBeVisible();
  await stale.close();
});
