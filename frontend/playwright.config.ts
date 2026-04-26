import { defineConfig, devices } from "@playwright/test";

function parsePort(name: string, fallback: number): number {
  const raw = process.env[name];
  const value = raw ? Number(raw) : fallback;
  if (!Number.isInteger(value) || value < 1 || value > 65_535) {
    throw new Error(`${name} must be a valid TCP port, got: ${raw ?? fallback}`);
  }
  return value;
}

const backendPort = parsePort("PLAYWRIGHT_BACKEND_PORT", 8000);
const frontendPort = parsePort("PLAYWRIGHT_FRONTEND_PORT", 4173);

if (backendPort === frontendPort) {
  throw new Error("PLAYWRIGHT_BACKEND_PORT and PLAYWRIGHT_FRONTEND_PORT must be different.");
}

const backendOrigin = `http://127.0.0.1:${backendPort}`;
const frontendOrigin = `http://127.0.0.1:${frontendPort}`;

export default defineConfig({
  testDir: "./e2e",
  timeout: 30_000,
  expect: {
    timeout: 5_000,
  },
  use: {
    baseURL: frontendOrigin,
    trace: "on-first-retry",
    screenshot: "only-on-failure",
    viewport: { width: 1440, height: 1200 },
  },
  projects: [
    {
      name: "chromium",
      use: { ...devices["Desktop Chrome"] },
    },
  ],
  webServer: [
    {
      command: `..\\.venv\\Scripts\\python.exe -m uvicorn a3presentation.main:app --app-dir ..\\src --host 127.0.0.1 --port ${backendPort}`,
      port: backendPort,
      reuseExistingServer: true,
      timeout: 30_000,
    },
    {
      command: `yarn dev --host 127.0.0.1 --port ${frontendPort}`,
      env: {
        ...process.env,
        VITE_PROXY_API_TARGET: backendOrigin,
      },
      port: frontendPort,
      reuseExistingServer: true,
      timeout: 30_000,
    },
  ],
});
