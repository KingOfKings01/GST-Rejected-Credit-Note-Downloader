/**
 * Detects the .dimmer-holder loader element on the page and waits until it disappears stably.
 * 
 * @param {import('playwright').Page} pageInstance - The Playwright Page instance.
 * @param {number} quietPeriodMs - The continuous time in milliseconds that the dimmer must remain hidden (default: 300).
 * @param {number} checkIntervalMs - The checking interval in milliseconds (default: 200).
 */
async function waitForDimmer(pageInstance, quietPeriodMs = 300, checkIntervalMs = 200) {
  console.log('Waiting for dimmer loader to disappear...');
  
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

  console.log(`Verifying that dimmer does not reappear for ${quietPeriodMs / 1000} seconds...`);

  while (elapsedMs < quietPeriodMs) {
    await pageInstance.waitForTimeout(checkIntervalMs);
    elapsedMs += checkIntervalMs;

    if (await isDimmerVisible()) {
      console.log('Dimmer reappeared! Resetting timer and waiting for it to disappear again...');
      
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
      console.log(`Dimmer cleared. Restarting the ${quietPeriodMs / 1000}-second confirmation window...`);
    }
  }

  console.log('Dimmer loader disappeared and remained stable ✅');
}

module.exports = {
  waitForDimmer
};
