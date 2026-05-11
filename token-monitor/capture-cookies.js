const puppeteer = require("puppeteer-core");
const fs = require("fs");
const path = require("path");

const CHROME_PATH = String.raw`C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe`;
const MIMO_URL = "https://platform.xiaomimimo.com/console/plan-manage";
const STORE_PATH = path.join(__dirname, "store.json");
const ENV_PATH = path.join(__dirname, "..", ".env");

async function main() {
  console.log("正在启动 Chrome...");
  const browser = await puppeteer.launch({
    executablePath: CHROME_PATH,
    headless: false,
    defaultViewport: { width: 1200, height: 800 },
    args: ["--no-first-run"],
  });

  const page = await browser.newPage();
  console.log("正在打开 MiMo 平台...");
  await page.goto(MIMO_URL, { waitUntil: "networkidle2", timeout: 30000 });

  console.log("\n========================================");
  console.log("  请在弹出的 Chrome 窗口中登录小米账号");
  console.log("  登录成功后，页面会自动检测并保存 Cookie");
  console.log("  请不要关闭此窗口");
  console.log("========================================\n");

  // Poll for cookies every 3 seconds
  let saved = false;
  const checkCookie = async () => {
    if (saved) return;
    try {
      const cookies = await page.cookies("https://platform.xiaomimimo.com");
      const cookieStr = cookies.map((c) => `${c.name}=${c.value}`).join(";");

      // Check if we have the key cookies
      const hasToken = cookies.some((c) => c.name === "api-platform_serviceToken");
      const hasSlh = cookies.some((c) => c.name === "api-platform_slh");

      if (hasToken && hasSlh) {
        saved = true;
        console.log("Cookie 获取成功！正在保存...");

        // Save to store.json
        let store = {};
        try { store = JSON.parse(fs.readFileSync(STORE_PATH, "utf-8")); } catch {}
        store.cookies = cookieStr;
        fs.writeFileSync(STORE_PATH, JSON.stringify(store, null, 2));

        // Also save to .env
        let env = "";
        try { env = fs.readFileSync(ENV_PATH, "utf-8"); } catch {}
        if (env.includes("MIMO_CONSOLE_COOKIES=")) {
          env = env.replace(/MIMO_CONSOLE_COOKIES=.*/, `MIMO_CONSOLE_COOKIES=${cookieStr}`);
        } else {
          env += `\n# Render Deploy Hook\nMIMO_CONSOLE_COOKIES=${cookieStr}`;
        }
        // Don't overwrite .env if it has other content we don't want to break
        // Just save to store.json as primary
        console.log("Cookie 已保存到 store.json");
        console.log("现在可以刷新监控页面查看数据了！");
        console.log("关闭此窗口和 Chrome 即可。");

        // Keep browser open for 5 seconds then close
        setTimeout(() => browser.close(), 5000);
      }
    } catch (e) {
      // Page might have navigated or errored
    }
  };

  // Check every 3 seconds
  const interval = setInterval(checkCookie, 3000);

  // Also check on page load/navigation
  page.on("load", checkCookie);

  // Timeout after 5 minutes
  setTimeout(() => {
    if (!saved) {
      console.log("\n超时未检测到登录 Cookie，请重新运行此脚本。");
      clearInterval(interval);
      browser.close();
    }
  }, 5 * 60 * 1000);
}

main().catch((e) => {
  console.error("错误:", e.message);
  process.exit(1);
});
