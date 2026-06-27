import { Page, chromium } from "@playwright/test";
import fs from "fs";
import path from "path";
const XlsxPopulate = require('xlsx-populate');
import { ChildProcess, spawn } from "child_process";

const FIXED_COUNT_OF_SHEET_ON_TEMPLATE = 13; // จำนวน sheet ที่อยู่ใน template.xlsx ที่ไม่ต้อง copy

const ALL_STYLE_NAMES = [
  "bold",
  "italic",
  "underline",
  "strikethrough",
  "subscript",
  "superscript",
  "fontSize",
  "fontFamily",
  "fontGenericFamily",
  "fontScheme",
  "fontColor",
  "fill",
  "border",
  "borderColor",
  "borderStyle",
  "diagonalBorderDirection",
  "diagonalBorderStyle",
  "diagonalBorderColor",
  "horizontalAlignment",
  "justifyLastLine",
  "indent",
  "verticalAlignment",
  "wrapText",
  "shrinkToFit",
  "textDirection",
  "textRotation",
  "angleTextCounterclockwise",
  "angleTextClockwise",
  "rotateTextUp",
  "rotateTextDown",
  "verticalText",
  "numberFormat",
];

function copyWithTimestamp(sourcePath: string, destFolder: string) {
  const now = new Date();
  const timestamp = now
    .toLocaleString("th-TH", {
      dateStyle: "short",
      timeStyle: "medium",
    })
    .replace(/\//g, "-")
    .replace(/:/g, "");

  const ext = path.extname(sourcePath);
  const fileName = `${timestamp}${ext}`;

  fs.mkdirSync(destFolder, { recursive: true });

  const destPath = path.join(destFolder, fileName);

  fs.copyFileSync(sourcePath, destPath);
  console.log(`Copied to: ${destPath}`);
  return fileName;
}

async function getDataFromJira({ page, chrome }: { page: Page, chrome: ChildProcess }) {
  await page.goto("https://cy-autoxjira.atlassian.net/browse/TLMSIT-96507"); // เปลี่ยน URL ให้ตรงกับ Jira ที่ต้องการ

  await page.waitForTimeout(15000); // รอให้หน้าโหลดข้อมูล

  const myIframe = page.frameLocator('iframe[id^="com.thed.zephyr.je__viewissue-teststep-issuecontent-bdd-two"]');
  const rows = myIframe.locator("#ISSUEVIEW_TESTSTEP div.zs-body-container.overflow-y.zs-scroll-grid > div")
  const rowCount = await rows.count() - 1;
  console.log("rows count", rowCount);

  if (rowCount <= 0) {
    console.log("No data found in Jira test steps.");
    chrome.kill();
    return [];
  }

  const data: (string | number)[][] = [];

  for (let i = 0; i < rowCount; i++) {
    const row = rows.nth(i);
    const cells = row.locator("div.zs-grid-body-wrapper > div"); // ข้อมูล (td)
    const cellCount = 3;

    const rowData: (string | number)[] = [];
    for (let j = 0; j < cellCount; j++) {
      if (j === 0) {
        rowData.push(i + 1); // เพิ่มลำดับที่
      } else {
        const text = await cells.nth(j + 1).innerText();
        rowData.push(text.trim());
      }
    }
    data.push(rowData);
  }

  console.log("data", data);
  return data;
}

function copySheet({
  workbook,
  numberSheet,
}: {
  workbook: any;
  numberSheet: number;
}) {
  const sourceSheet = workbook.sheet(2);
  const newSheet = workbook.cloneSheet(
    sourceSheet,
    numberSheet.toString(),
    numberSheet + 1,
  );
  newSheet.cell(2, 2).formula(`=TC!B${numberSheet + 4}`);
  newSheet.cell(2, 3).formula(`=TC!C${numberSheet + 4}`);
}

function copyRowWithStyle({
  sheet,
  sourceRowNum,
  destRowNum,
  maxColumns = 3,
}: {
  sheet: any;
  sourceRowNum: number;
  destRowNum: number;
  maxColumns?: number;
}) {
  const sourceRow = sheet.row(sourceRowNum);
  const destRow = sheet.row(destRowNum);

  // 1. Copy the row height if it is set
  if (sourceRow.height()) {
    destRow.height(sourceRow.height());
  }

  // 2. Loop through each cell to transfer values and styling
  for (let colIdx = 1; colIdx <= maxColumns; colIdx++) {
    const sourceCell = sheet.cell(sourceRowNum, colIdx);
    const destCell = sheet.cell(destRowNum, colIdx);

    // Copy value
    destCell.value(sourceCell.value());

    // Extract style properties and apply them to the destination cell
    const currentStyles = sourceCell.style(ALL_STYLE_NAMES);
    destCell.style(currentStyles);
  }
}

async function writeDataToExcel(filePath: string, data: (string | number)[][]) {
  const workbook = await XlsxPopulate.fromFileAsync(filePath);

  if (data.length > FIXED_COUNT_OF_SHEET_ON_TEMPLATE) {
    for (let i = FIXED_COUNT_OF_SHEET_ON_TEMPLATE; i < data.length; i++) {
      copySheet({ workbook, numberSheet: i + 1 });
    }
  }

  const sheet = workbook.sheet(1); // หรือ workbook.sheet("ชื่อ sheet")

  data.forEach((row, rowIndex) => {
    row.forEach((value, colIndex) => {
      sheet.cell(rowIndex + 5, colIndex + 1).value(value);
      if (rowIndex < data.length - 1) {
        copyRowWithStyle({
          sheet,
          sourceRowNum: rowIndex + 5, // Assuming the template starts at row 5
          destRowNum: rowIndex + 6, // Copy to the next row
        });
      }
    });
  });

  await workbook.toFileAsync(filePath);
}

async function main() {
  const chrome = spawn(
    "C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe",
    [
      "--remote-debugging-port=9222",
      `--user-data-dir=C:\\temp\\playwright-profile`,
    ],
    { detached: true, stdio: "ignore" }
  );

  // รอ Chrome เปิด
  await new Promise((resolve) => setTimeout(resolve, 2000));
  const browser = await chromium.connectOverCDP("http://127.0.0.1:9222");

  const context = browser.contexts()[0]; // Profile แรกที่เปิดอยู่
  const page = context.pages()[0]; // แท็บแรกที่เปิดอยู่

  // คัดลอกไฟล์ Template.xlsx ไปยังโฟลเดอร์ test-steps พร้อมกับ timestamp
  const resultFileName = await copyWithTimestamp(
    "src/Template.xlsx",
    "src/test-steps",
  );

  // ดึงข้อมูลจาก Jira
  const data = await getDataFromJira({ page, chrome });

  // เขียนข้อมูลลงในไฟล์ Excel
  await writeDataToExcel(`src/test-steps/${resultFileName}`, data);

  chrome.kill(); // ปิด Chrome หลังจากทำงานเสร็จ
}

main();
