
async function retry(fn, retries = 2, delay = 1000) {
    for (let i = 0; i < retries; i++) {
        try {
            return await fn();
        } catch (e) {
            if (i === retries - 1) throw e;
            await new Promise(r => setTimeout(r, delay));
        }
    }
}

/**
 * Executes the login loop sequence on a shared page context instance.
 * @param {import('playwright').Page} page - The active page instance provided by the loop runner
 * @param {string} username 
 * @param {string} password 
 * @returns {Promise<boolean>} True if login is verified, false otherwise
 */
async function loginGST(page, username, password, browserZoom = '0.85') {
    try {
        await retry(async () => {
            await page.goto('https://services.gst.gov.in/services/login');
        });
        // Zoom out according to configured browserZoom value
        await page.evaluate((zoom) => {
            const scaleVal = parseFloat(zoom) || 0.85;
            document.body.style.transform = `scale(${scaleVal})`;
            document.body.style.transformOrigin = 'top left'; // Keeps content aligned to the top-left
            document.body.style.width = `${(1 / scaleVal) * 100}%`; // Compenses for the scale down so layout doesn't break
        }, browserZoom);


        let loginSuccessful = false;
        let skipRequested = false;

        try {
            await page.exposeFunction('requestSkipClient', () => {
                skipRequested = true;
            });
        } catch (exposeErr) {
            // Safe to ignore if function already exposed
        }

        while (!loginSuccessful) {
            // Re-bind the Escape keypress listener inside the browser context
            await page.evaluate(() => {
                if (!window.__hasEscapeSkipListener) {
                    window.__hasEscapeSkipListener = true;
                    window.addEventListener('keydown', (e) => {
                        if (e.key === 'Escape') {
                            window.requestSkipClient();
                        }
                    });
                }
            }).catch(() => {});

            await page.locator('#username').fill(username);
            await page.locator('#user_pass').fill(password);
            
            await page.keyboard.press('Tab');

            console.log("\n⚠️ ACTION REQUIRED: Please type the Captcha into the browser and press ENTER.");
            console.log("💡 Tip: You can press the ESCAPE key inside the browser window at any time to skip this client.");

            let checkInterval;
            const skipPromise = new Promise((resolve) => {
                checkInterval = setInterval(() => {
                    if (skipRequested) {
                        clearInterval(checkInterval);
                        resolve('skip');
                    }
                }, 300);
            });

            // Wait for authentication states or user skip request
            const result = await Promise.race([
                page.waitForURL('https://services.gst.gov.in/services/auth/fowelcome', { timeout: 0 }).then(() => 'success'),
                page.waitForSelector('.alert.alert-danger:has-text("Invalid Username or Password")', { timeout: 0 }).then(() => 'creds_error'),
                page.waitForSelector('span[data-ng-bind="trans.ERR_SWEB_9000"]', { timeout: 0 }).then(() => 'captcha_error'),
                skipPromise
            ]);

            clearInterval(checkInterval);

            if (result === 'skip') {
                console.log(`\n⏭️ User requested to skip client: ${username}`);
                return 'skipped';
            }

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