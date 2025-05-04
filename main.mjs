#!/usr/bin/env node

import os from "os";
import fs from "fs";
import path from "path";
import { chromium } from "playwright";

// --- Configuration ---
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

// --- Arguments ---
const workspaceName = process.argv[2];

if (!CHROME_PATH || !fs.existsSync(CHROME_PATH)) {
  console.error("❌ Chrome not found in the default location for your OS.");
  process.exit(1);
}

if (!workspaceName) {
  console.error("❌ Usage: script <workspace_name>");
  process.exit(1);
}

// --- Browser Functions ---
async function launchChrome() {
  console.log("🚀 Launching Chrome...");
  const context = await chromium.launchPersistentContext(USER_DATA_DIR, {
    headless: false,
    executablePath: CHROME_PATH,
    args: [
      "--profile-directory=Default",
      "--disable-popup-blocking",
      "--kiosk-printing",
    ],
  });
  console.log("✅ Chrome launched.");
  return context;
}

async function openWorkspace(page) {
  const selector = `div[aria-label*='${workspaceName}']`;
  console.log(`🌐 Navigating to ${LOOP_URL}...`);
  await page.goto(LOOP_URL);

  console.log(`🔍 Looking for workspace '${workspaceName}'...`);
  await page.waitForSelector(selector);
  const button = await page.$(selector);

  if (!button) {
    throw new Error(`❌ Workspace '${workspaceName}' not found.`);
  }

  console.log(`📂 Opening workspace '${workspaceName}'...`);
  await button.click();
  console.log("✅ Workspace opened.");
}

async function getTreeItems(page) {
  console.log("📋 Waiting for tree items to load...");
  await page.waitForSelector("*[role='treeitem']");
  await page.waitForSelector('[data-testid^="PageSubtree_"]');
  await page.waitForTimeout(5000);
  const items = await page.$$("*[role='treeitem']");
  console.log(`📦 Found ${items.length} tree item(s).`);
  return items;
}

async function exportItemToPDF(context, page, item, index) {
  console.log(`➡ Clicking item ${index + 1}...`);
  await item.click();

  try {
    console.log("⏳ Waiting for 'Más opciones' button...");
    await page.waitForSelector("button[aria-label='Más opciones']", { timeout: 10000 });
    await page.click("button[aria-label='Más opciones']");
    await page.waitForSelector('div[role="menuitem"]');
    console.log("📑 'Más opciones' menu opened.");
  } catch {
    console.warn(`⚠ Skipping item ${index + 1}: Options menu not available.`);
    return "skipped";
  }

  const exportButton = await findPDFOption(page);
  if (!exportButton) {
    console.warn(`⚠ Skipping item ${index + 1}: No PDF export option found.`);
    return "skipped";
  }

  console.log("🖨 Clicking PDF export option...");
  const [popup] = await Promise.all([
    context.waitForEvent("page"),
    exportButton.click(),
  ]);

  try {
    console.log("🪟 Handling print popup...");
    await handlePopupPrint(popup, index + 1);
    await popup.close();
    console.log(`✔ Exported item ${index + 1}`);
    return "exported";
  } catch (error) {
    console.error(`❌ Error exporting item ${index + 1}:`, error);
    return "skipped";
  }
}

async function findPDFOption(page) {
  const menuItems = await page.$$('div[role="menuitem"]');
  for (const item of menuItems) {
    const text = await item.evaluate(el => el.textContent?.trim() || "");
    if (text.includes("PDF")) return item;
  }
  return null;
}

async function handlePopupPrint(popup, index) {
  await popup.waitForLoadState();

  await popup.evaluate(() => {
    window.__printed = false;
    const originalPrint = window.print;
    window.print = function () {
      window.__printed = true;
      originalPrint.call(window);
    };
  });

  try {
    await popup.waitForFunction(() => window.__printed, {}, { timeout: 10000 });
    console.log(`🖨 Print triggered for item ${index}`);
  } catch {
    console.warn(`⚠ Print not triggered for item ${index}`);
  }
}

// --- Tree Expansion ---
const expandedItemsKeys = new Set();

async function expandTreeItems(page) {
  console.log("🌲 Expanding tree items...");
  await page.waitForSelector(".fui-TreeItemLayout__expandIcon", { timeout: 10000 });
  const expandIcons = await page.$$(".fui-TreeItemLayout__expandIcon");
  let expanded = false;

  for (const icon of expandIcons) {
    const label = await icon.evaluate(el => el.parentElement?.textContent?.trim() || "");
    if (expandedItemsKeys.has(label)) continue;

    console.log(`➕ Expanding: ${label}`);
    try {
      await icon.click();
      await page.waitForTimeout(500);
    } catch (err) {
      console.error(`❌ Error expanding '${label}':`, err);
    }

    expandedItemsKeys.add(label);
    expanded = true;
  }

  if (expanded) {
    console.log("🔁 Expanding nested items...");
    await expandTreeItems(page);
  } else {
    console.log("✅ All tree items expanded.");
  }
}

// --- Main ---
async function run() {
  const context = await launchChrome();
  let successCount = 0;
  let skippedCount = 0;

  try {
    const page = await context.newPage();
    console.log("📄 New page created.");

    await openWorkspace(page);
    await expandTreeItems(page);
    const items = await getTreeItems(page);

    for (let i = 0; i < items.length; i++) {
      console.log(`\n=== 📦 Processing item ${i + 1} of ${items.length} ===`);
      const result = await exportItemToPDF(context, page, items[i], i);
      if (result === "exported") successCount++;
      else skippedCount++;
      await page.waitForTimeout(1000);
    }

    console.log("\n🎉 All items processed.");
    console.log(`✔ Successfully exported: ${successCount}`);
    console.log(`⚠ Skipped: ${skippedCount}`);
  } catch (err) {
    console.error("❌ Unexpected error:", err);
    process.exit(1);
  } finally {
    await context.close();
    console.log("🛑 Browser context closed.");
  }
}

// --- Start ---
await run();
console.log("🚀 Script finished.");
