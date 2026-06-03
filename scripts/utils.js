/**
 * Detects the .dimmer-holder loader element on the page and waits until it disappears stably.
 * 
 * @param {import('playwright').Page} pageInstance - The Playwright Page instance.
 * @param {number} quietPeriodMs - The continuous time in milliseconds that the dimmer must remain hidden (default: 300).
 * @param {number} checkIntervalMs - The checking interval in milliseconds (default: 200).
 */
async function waitForDimmer(pageInstance, quietPeriodMs = 300, checkIntervalMs = 200) {
  
  // Helper function to check if the dimmer is currently visible
  const isDimmerVisible = async () => {
    return await pageInstance.evaluate(() => {
      const dimmer = document.querySelector('.dimmer-holder');
      if (!dimmer) return false;
      const style = window.getComputedStyle(dimmer);
      return (
        style.display !== 'none' &&
        style.visibility !== 'hidden' &&
        parseFloat(style.opacity) > 0
      );
    });
  };

  // 1. Initial wait for dimmer to disappear
  await pageInstance.waitForFunction(() => {
    const dimmer = document.querySelector('.dimmer-holder');
    if (!dimmer) return true;
    const style = window.getComputedStyle(dimmer);
    return (
      style.display === 'none' || 
      style.visibility === 'hidden' || 
      parseFloat(style.opacity) === 0
    );
  }, { timeout: 30000 }).catch(() => {});

  // 2. Verify that it stays gone (quiet period)
  let elapsedMs = 0;
  while (elapsedMs < quietPeriodMs) {
    await pageInstance.waitForTimeout(checkIntervalMs);
    elapsedMs += checkIntervalMs;

    if (await isDimmerVisible()) {    
      // Wait for it to disappear again
      await pageInstance.waitForFunction(() => {
        const dimmer = document.querySelector('.dimmer-holder');
        if (!dimmer) return true;
        const style = window.getComputedStyle(dimmer);
        return (
          style.display === 'none' || 
          style.visibility === 'hidden' || 
          parseFloat(style.opacity) === 0
        );
      }, { timeout: 30000 });

      // Reset elapsed time to restart the confirmation window
      elapsedMs = 0;
    }
  }
}

module.exports = {
  waitForDimmer
};
