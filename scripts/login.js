
/**
 * Executes the login loop sequence on a shared page context instance.
 * @param {import('playwright').Page} page - The active page instance provided by the loop runner
 * @param {string} username 
 * @param {string} password 
 * @returns {Promise<boolean>} True if login is verified, false otherwise
 */
async function loginGST(page, username, password) {
    try {
        await page.goto('https://services.gst.gov.in/services/login');
        // Zoom out to 80% (equivalent to pressing Ctrl + Minus twice)
    await page.evaluate(() => {
        document.body.style.transform = 'scale(0.75)';
        document.body.style.transformOrigin = 'top left'; // Keeps content aligned to the top-left
        document.body.style.width = '135%'; // Compenses for the scale down so layout doesn't break
    });


        let loginSuccessful = false;

        while (!loginSuccessful) {
            await page.locator('#username').fill(username);
            await page.locator('#user_pass').fill(password);
            
            await page.keyboard.press('Tab');

            console.error("\n⚠️ ACTION REQUIRED: Please type the Captcha into the browser and press ENTER.");

            // Wait for authentication states
            await Promise.race([
                page.waitForURL('https://services.gst.gov.in/services/auth/fowelcome', { timeout: 0 }),
                page.waitForSelector('.alert.alert-danger:has-text("Invalid Username or Password")', { timeout: 0 }),
                page.waitForSelector('span[data-ng-bind="trans.ERR_SWEB_9000"]', { timeout: 0 })
            ]);

            const currentUrl = page.url();
            const isCredsErrorVisible = await page.locator('.alert.alert-danger:has-text("Invalid Username or Password")').isVisible();
            const isCaptchaErrorVisible = await page.locator('span[data-ng-bind="trans.ERR_SWEB_9000"]').isVisible();

            if (currentUrl.includes('/auth/fowelcome')) {
                loginSuccessful = true;
                return true; // Signal back success status code
            } 
            else if (isCredsErrorVisible) {
                console.error("❌ Error: Invalid Username or Password. Skipping process for this client.");
                return false; 
            } 
            else if (isCaptchaErrorVisible) {
                await page.waitForTimeout(1000); 
            }
        }
    } catch (error) {
        console.error("An error occurred during login routing operation:", error);
        return false;
    }
}

module.exports = { loginGST };