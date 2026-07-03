
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
 * @param {string} browserZoom
 * @param {Map} captchaResolvers
 * @returns {Promise<boolean|string>} True if login is verified, 'skipped' if skipped, false otherwise
 */
async function loginGST(page, username, password, browserZoom = '0.85', captchaResolvers = null) {
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
            // Re-bind the Escape keypress listener inside the browser context (only if headed)
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
            
            // Take screenshot of #imgCaptcha
            const captchaElement = page.locator('#imgCaptcha');
            await captchaElement.waitFor({ state: 'visible', timeout: 15000 });
            
            // Give captcha a brief moment to load completely
            await page.waitForTimeout(500);
            const captchaBase64 = await captchaElement.screenshot({ type: 'png' }).then(buf => buf.toString('base64'));

            // Emit captcha request to Electron UI
            console.log(`CLIENT_PROGRESS:CAPTCHA_REQUEST:${username}:${captchaBase64}`);

            let checkInterval;
            const skipPromise = new Promise((resolve) => {
                checkInterval = setInterval(() => {
                    if (skipRequested) {
                        clearInterval(checkInterval);
                        resolve({ action: 'skip' });
                    }
                }, 300);
            });

            // Wait for user input or refresh from IPC (via the stdin router) or physical browser ESC skip
            let response = null;
            if (captchaResolvers) {
                const uiResponsePromise = new Promise((resolve) => {
                    captchaResolvers.set(username, resolve);
                });
                response = await Promise.race([uiResponsePromise, skipPromise]);
            } else {
                response = await skipPromise;
            }

            clearInterval(checkInterval);

            if (response.action === 'skip') {
                console.log(`\n⏭️ User requested to skip client: ${username}`);
                console.log(`CLIENT_PROGRESS:CAPTCHA_SOLVED:${username}`); // Remove from UI queue
                return 'skipped';
            }

            if (response.action === 'refresh') {
                console.log(`System: Refreshing captcha for ${username}`);
                try {
                    const refreshBtn = page.locator('.captcha-refresh, a[ng-click*="refresh"], a[href*="refresh"], .glyphicon-refresh').first();
                    await refreshBtn.waitFor({ state: 'visible', timeout: 5000 });
                    await refreshBtn.click({ timeout: 5000 });
                } catch (clickErr) {
                    console.log(`System: Refresh button click timed out or failed. Reloading page as fallback...`);
                    await page.reload({ waitUntil: 'networkidle', timeout: 15000 }).catch(() => {});
                }
                await page.waitForTimeout(1000);
                continue; // Loop back and capture new captcha
            }

            // Input captcha text and click login
            const captchaText = response.text;
            await page.locator('#captcha').fill(captchaText);
            
            const loginBtn = page.locator('button[type="submit"], button:has-text("Login")').first();
            await loginBtn.click();

            // Wait for authentication states or errors
            const result = await Promise.race([
                page.waitForURL('https://services.gst.gov.in/services/auth/fowelcome', { timeout: 12000 }).then(() => 'success'),
                page.waitForSelector('.alert.alert-danger:has-text("Invalid Username or Password")', { timeout: 6000 }).then(() => 'creds_error'),
                page.waitForSelector('span[data-ng-bind="trans.ERR_SWEB_9000"]', { timeout: 6000 }).then(() => 'captcha_error'),
                page.waitForTimeout(10000).then(() => 'timeout')
            ]);

            if (result === 'success') {
                loginSuccessful = true;
                console.log(`CLIENT_PROGRESS:CAPTCHA_SOLVED:${username}`);
                return true; // Signal back success status code
            } 
            else if (result === 'creds_error') {
                console.error(`❌ Error: Invalid Username or Password for ${username}. Skipping.`);
                console.log(`CLIENT_PROGRESS:CAPTCHA_SOLVED:${username}`); // Remove from UI queue
                return false; 
            } 
            else {
                // Captcha error or timeout
                console.log(`⚠️ Captcha validation failed or timed out for ${username}. Retrying...`);
                await page.waitForTimeout(1000); 
            }
        }
    } catch (error) {
        console.error("An error occurred during login routing operation:", error);
        if (captchaResolvers) captchaResolvers.delete(username);
        return false;
    }
}

module.exports = { loginGST };