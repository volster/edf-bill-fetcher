import { defineConfig, devices } from "@playwright/test";

const browser = process.env.EDF_BROWSER || "chromium";
export default defineConfig({
  testDir: "./tests",
  fullyParallel: false,
  workers: 1,
  timeout: 90_000,
  expect: { timeout: 20_000 },
  use: { baseURL: "http://127.0.0.1:8766", trace: "retain-on-failure" },
  projects: [
    {
      name: browser,
      use: {
        ...devices[
          browser === "webkit"
            ? "Desktop Safari"
            : browser === "firefox"
              ? "Desktop Firefox"
              : "Desktop Chrome"
        ],
      },
    },
  ],
  webServer: {
    command: `${process.env.EDF_TEST_PYTHON || "python"} -m tests.web_server`,
    url: "http://127.0.0.1:8766",
    reuseExistingServer: false,
    timeout: 60_000,
    cwd: "..",
  },
});
