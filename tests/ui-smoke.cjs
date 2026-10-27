// Run against the local static server with an existing Playwright installation.
// All fixtures are synthetic and stay in memory. No AI request is made.
const { chromium } = require('playwright');
const assert = require('node:assert/strict');
const TARGET_URL = process.env.TARGET_URL || 'http://127.0.0.1:4180/';
const libraryPath = process.env.XLSX_TEST_SCRIPT;
const csv = '學校名稱,系所名稱,關鍵字\n測試大學,財務金融學系,finance;金融\n測試大學,會計學系,accounting;會計';

async function downloadRows(page, selector) {
  const pending = page.waitForEvent('download');
  await page.click(selector);
  const download = await pending;
  const stream = await download.createReadStream();
  const chunks = [];
  for await (const chunk of stream) chunks.push(chunk);
  const base64 = Buffer.concat(chunks).toString('base64');
  return page.evaluate((data) => {
    const workbook = XLSX.read(data, { type: 'base64' });
    return {
      names: workbook.SheetNames,
      rows: XLSX.utils.sheet_to_json(workbook.Sheets[workbook.SheetNames[0]], { header: 1, defval: '', raw: false }),
    };
  }, base64);
}

(async () => {
  const browser = await chromium.launch({ headless: process.env.HEADLESS === '1', channel: process.env.BROWSER_CHANNEL || 'chrome' });
  try {
    const context = await browser.newContext({ viewport: { width: 1440, height: 1000 }, acceptDownloads: true, permissions: ['clipboard-read', 'clipboard-write'] });
    if (libraryPath) {
      await context.route('https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js', (route) => route.fulfill({ path: libraryPath, contentType: 'application/javascript' }));
    }
    const page = await context.newPage();
    const errors = [];
    page.on('pageerror', (error) => errors.push(error.message));
    await page.goto(TARGET_URL);
    await page.waitForFunction(() => Boolean(window.XLSX));
    const fixtures = await page.evaluate(() => {
      const build = (rows) => {
        const workbook = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet(rows), 'Books');
        return XLSX.write(workbook, { type: 'base64', bookType: 'xlsx' });
      };
      return {
        source: build([
          ['Book Title', 'ISBN', 'Subject', 'Author'],
          ['Corporate Finance', '978-1-111-23456-7', 'Business', 'Alex Chen'],
          ['Accounting', '9781111234568', 'Business', 'Jordan Lin'],
          ['Clean Energy', '9781111234569', 'Engineering', 'Sam Wang'],
        ]),
        purchased: build([['ISBN', 'Title'], ['9781111234567', 'Corporate Finance']]),
      };
    });
    for (const mode of ['department', 'purchase', 'upload', 'combined']) {
      await page.goto(TARGET_URL);
      await page.reload();
      await page.waitForFunction(() => Boolean(window.XLSX));
      await page.locator('#home:visible').waitFor();
      await page.click(`[data-mode="${mode}"]`);
      await page.click('#nextButton');
      assert(await page.locator('#stepError').isVisible(), 'Empty upload must be blocked');
      await page.setInputFiles('#sourceFile', { name: 'books.xlsx', mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', buffer: Buffer.from(fixtures.source, 'base64') });
      await page.locator('#sourceFileInfo:visible').waitFor();
      if (['purchase', 'combined'].includes(mode)) {
        await page.setInputFiles('#purchasedFile', { name: 'purchased.xlsx', mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', buffer: Buffer.from(fixtures.purchased, 'base64') });
        await page.locator('[data-purchased-field="purchased_key"]').waitFor({ state: 'attached' });
      }
      await page.click('#nextButton');
      assert.equal(await page.locator('.mapping-row:visible').count(), mode === 'combined' ? 4 : 2);
      if (mode !== 'purchase') {
        await page.selectOption('[data-source-field="title"]', '');
        await page.click('#nextButton');
        assert(await page.locator('[data-source-field="title"]').getAttribute('aria-invalid'));
        await page.selectOption('[data-source-field="title"]', '0');
        assert(await page.locator('[data-source-field="subject_main"]').evaluate(el => el.required));
        await page.selectOption('[data-source-field="subject_main"]', '');
        await page.click('#nextButton');
        assert(await page.locator('[data-source-field="subject_main"]').getAttribute('aria-invalid'));
        assert(await page.locator('[data-page="mapping"]').isVisible());
        await page.selectOption('[data-source-field="subject_main"]', '2');
      } else {
        assert.equal(await page.locator('[data-source-field="subject_main"]').evaluate(el => el.required), false);
      }
      await page.click('#nextButton');
      if (['department', 'combined'].includes(mode)) {
        if (mode === 'department') {
          assert.equal(await page.locator('#sourceDepartmentText').count(), 0);
          assert.equal((await page.locator('#copyPromptButton').innerText()).trim(), '下一步');
          await page.click('#copyPromptButton');
          assert(await page.locator('#schoolName').getAttribute('aria-invalid'));
          await page.fill('#schoolName', '測試大學');
          await page.click('#copyPromptButton');
          await page.getByRole('heading', { name: 'AI 指令已複製', exact: true }).waitFor();
          const prompt = await page.evaluate(() => navigator.clipboard.readText());
          assert(prompt.includes('測試大學') && prompt.includes('官方網站') && prompt.includes('自行查找'));
          assert(!prompt.includes('Corporate Finance'), 'The AI prompt must not include the uploaded workbook');
          assert.equal(await page.locator('#copyPromptConfirm').evaluate(el => el === document.activeElement), true);
          await page.keyboard.press('Escape');
          assert(await page.locator('#copyPromptDialog').isVisible(), 'Escape must not dismiss the acknowledgement');
          await page.mouse.click(5, 5);
          assert(await page.locator('#copyPromptDialog').isVisible(), 'Backdrop click must not dismiss the acknowledgement');
          await page.click('#copyPromptConfirm');
          assert(await page.locator('#copyPromptDialog').isHidden());
          assert(await page.locator('[data-ai-page="1"]').isVisible());
          // Clipboard failure must offer manual copying and never claim success.
          await page.evaluate(() => {
            window.restoreClipboardWrite = navigator.clipboard.writeText.bind(navigator.clipboard);
            navigator.clipboard.writeText = async () => { throw new Error('Test: clipboard permission denied'); };
          });
          await page.click('[data-action="copy-again"]');
          await page.getByRole('heading', { name: '請手動複製 AI 指令', exact: true }).waitFor();
          assert((await page.locator('#copyPromptManual').inputValue()).includes('測試大學'));
          await page.click('#copyPromptConfirm');
          await page.evaluate(() => { navigator.clipboard.writeText = window.restoreClipboardWrite; delete window.restoreClipboardWrite; });
          await page.click('[data-action="return-csv"]');
        } else {
          await page.click('[data-action="existing-csv"]');
        }
        await page.fill('#schoolName', '測試大學');
        await page.fill('#keywordText', 'bad data');
        await page.click('[data-action="check-csv"]');
        assert(await page.locator('#keywordText').getAttribute('aria-invalid'));
        await page.fill('#keywordText', csv);
        await page.click('[data-action="check-csv"]');
        await page.locator('#csvFeedback .validation').waitFor();
        await page.click('#nextButton');
      }
      if (['upload', 'combined'].includes(mode)) {
        await page.fill('#schoolCode', 'TEST');
        await page.fill('#subjectCode', 'bad!');
        await page.click('#nextButton');
        assert(await page.locator('#subjectCode').getAttribute('aria-invalid'));
        await page.fill('#subjectCode', 'ebook');
        await page.click('#nextButton');
      }
      await page.click('#submitButton');
      await page.locator('[data-page="complete"]:visible').waitFor();
      if (mode !== 'upload') {
        const { rows } = await downloadRows(page, '#downloadOutput');
        assert.equal(rows.length, 4);
        assert.deepEqual(rows[0].slice(0, 4), ['Book Title', 'ISBN', 'Subject', 'Author']);
        if (['department', 'combined'].includes(mode)) assert.equal(rows[1][4], '財務金融學系');
        if (['purchase', 'combined'].includes(mode)) {
          assert.equal(rows[1][7], '已購買');
          assert.equal(rows[2][7], '未購買');
        }
      }
      if (['upload', 'combined'].includes(mode)) {
        const { names, rows } = await downloadRows(page, '#downloadUpload');
        assert.deepEqual(names, ['Tablib Dataset']);
        assert.equal(rows.length, 4);
        assert.equal(rows[1][rows[0].indexOf('主題代碼')], 'EBOOK');
        assert.equal(rows[1][rows[0].indexOf('題名')], 'Corporate Finance');
      }
      await page.click('#backButton');
      await page.click('[data-step="mapping"]');
      const changedField = mode === 'purchase' ? '[data-source-field="source_purchase_key"]' : '[data-source-field="title"]';
      await page.selectOption(changedField, '');
      assert.equal(await page.locator('#downloadOutput').getAttribute('hidden'), '');
      assert.equal(await page.locator('#downloadUpload').getAttribute('hidden'), '');
      await page.selectOption(changedField, mode === 'purchase' ? '1' : '0');
      for (const width of [375, 768, 1024, 1440]) {
        await page.setViewportSize({ width, height: 1000 });
        assert(await page.evaluate(() => document.documentElement.scrollWidth <= innerWidth), `Page overflows at ${width}px`);
      }
      const modeBeforeSkip = await page.locator('#modeName').textContent();
      await page.locator('.skip').focus();
      await page.keyboard.press('Enter');
      assert.equal(await page.locator('#modeName').textContent(), modeBeforeSkip, 'Skip link must preserve the current workflow');
      await page.setViewportSize({ width: 1440, height: 1000 });
      console.log(`PASS: ${mode} validation, actual Excel download, and responsive layout`);
    }
    assert.deepEqual(errors, []);
  } finally {
    await browser.close();
  }
})().catch((error) => { console.error(error); process.exitCode = 1; });
