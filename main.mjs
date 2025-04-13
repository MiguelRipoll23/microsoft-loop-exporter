#!/usr/bin/env node

import os from "os";
import fs from "fs";
import path from "path";
import { chromium } from "playwright";

// Constants for paths and URLs
const CHROME_PATHS = {
  win32: "C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe",
  darwin: "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome",
  linux: "/usr/bin/google-chrome",
};

const USER_DATA_PATHS = {
  win32: path.join(os.homedir(), "AppData/Local/Google/Chrome/User Data"),
  darwin: path.join(os.homedir(), "Library/Application Support/Google/Chrome"),
  linux: path.join(os.homedir(), ".config/google-chrome"),
};

const CHROME_PATH = CHROME_PATHS[process.platform];
const USER_DATA_DIR = USER_DATA_PATHS[process.platform];
const LOOP_URL = "https://loop.microsoft.com";

if (!CHROME_PATH || !fs.existsSync(CHROME_PATH)) {
  console.error(
    "Chrome not found in the default location for your operating system."
  );
  process.exit(1);
}

// Get the workspace name from command-line argument
const workspaceName = process.argv[2];

if (!workspaceName) {
  console.error("Workspace name is required as an argument.");
  process.exit(1);
}

// Launches Chrome with persistent context
async function launchChrome() {
  return chromium.launchPersistentContext(USER_DATA_DIR, {
    headless: false,
    executablePath: CHROME_PATH,
    args: ["--profile-directory=Default", "--kiosk-printing"],
  });
}

// Opens the specified workspace within the loop
async function openWorkspace(page) {
  const selector = `div[aria-label*='${workspaceName}']`;
  await page.goto(LOOP_URL);
  await page.waitForSelector(selector, { timeout: 10000 });

  const button = await page.$(selector);
  if (!button) throw new Error("Workspace not found");
  await button.click();

  console.log("Workspace opened");
}

// Fetches all tree items from the workspace
async function getTreeItems(page) {
  await page.waitForSelector("div[role='treeitem']", { timeout: 10000 });
  const items = await page.$$("div[role='treeitem']");
  console.log(`Found ${items.length} items`);
  return items;
}

// Exports the selected item to PDF
async function exportItemToPDF(context, page, item, index) {
  console.log(`Exporting item ${index + 1}`);
  await item.click();

  try {
    await page.click("button[aria-label='Más opciones']", { timeout: 10000 });
    await page.waitForSelector('div[role="menuitem"]', { timeout: 10000 });
  } catch {
    console.warn(`Options menu not ready for item ${index + 1}`);
    return;
  }

  const exportButton = await findPDFOption(page);
  if (!exportButton) {
    console.warn(`Export option not found for item ${index + 1}`);
    return;
  }

  const [popup] = await Promise.all([
    context.waitForEvent("page"),
    exportButton.click(),
  ]);

  await handlePopupPrint(popup, index + 1);
  await popup.close();
}

// Finds the PDF export option in the menu
async function findPDFOption(page) {
  const items = await page.$$('div[role="menuitem"]');
  for (const item of items) {
    const text = await item.evaluate((el) => el.textContent?.trim() || "");
    if (text.includes("PDF")) return item;
  }
  return null;
}

// Handles the print action in the popup
async function handlePopupPrint(popup, index) {
  await popup.waitForLoadState();

  await popup.evaluate(() => {
    window.__printed = false;
    const originalPrint = window.print;
    window.print = () => {
      window.__printed = true;
      originalPrint.call(window);
    };
  });

  try {
    await popup.waitForFunction(() => window.__printed, {}, { timeout: 10000 });
    console.log(`Print triggered for item ${index}`);
  } catch {
    console.warn(`Print not triggered for item ${index}`);
  }
}

// Main execution function
async function run() {
  const context = await launchChrome();

  try {
    const page = await context.newPage();
    await openWorkspace(page);
    const items = await getTreeItems(page);

    for (let i = 0; i < items.length; i++) {
      await exportItemToPDF(context, page, items[i], i);
    }

    console.log("All items processed.");
  } catch (err) {
    console.error("Unexpected error:", err);
    process.exit(1);
  } finally {
    await context.close();
  }
}

// Start the process
await run();
