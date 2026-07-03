const { waitForDimmer } = require('./utils');

/**
 * Selects the Financial Year, Return Period, and Return Type on the IMS dashboard and clicks Search.
 * @param {import('playwright').Page} page - The active Playwright page instance
 * @param {Object} selections - The form selection parameters (Direct Text Strings)
 * @param {string} selections.financialYear - E.g., '2025-26'
 * @param {string} selections.returnPeriod - E.g., 'April'
 * @param {string} selections.returnType - E.g., 'GSTR-1/IFF'
 */
async function fillImsForm(page, selections) {
    // 1. Handle Financial Year Dropdown
    const finYearDropdown = page.locator('select[name="fin"]');
    await finYearDropdown.waitFor({ state: 'visible', timeout: 30000 });
    await finYearDropdown.selectOption({ label: selections.financialYear });

    // Wait for the website loader/dimmer to disappear after FY selection
    await waitForDimmer(page, 300);

    // 2. Handle Return Period Dropdown
    const periodDropdown = page.locator('select[name="mon"]');
    await periodDropdown.waitFor({ state: 'visible', timeout: 30000 });
    await periodDropdown.selectOption({ label: selections.returnPeriod });

    // 3. Handle Return Type Dropdown
    const returnTypeDropdown = page.locator('select[name="rtntyp"]');
    await returnTypeDropdown.waitFor({ state: 'visible', timeout: 30000 });
    await returnTypeDropdown.selectOption({ label: selections.returnType });

    // 4. Click the Search Button
    const searchButton = page.locator('button[data-ng-bind="trans.LBL_SCH"]');

    // Wait until the button is fully active and visible on screen
    await searchButton.waitFor({ state: 'visible', timeout: 30000 });

    await searchButton.click();

    // Wait for data results to finish rendering over the network
    await page.waitForLoadState('networkidle');
    
    // Wait for the loading mask/spinner to clear and the Angular data binding to finish rendering numbers
    await waitForDimmer(page, 300);
}

module.exports = { fillImsForm };