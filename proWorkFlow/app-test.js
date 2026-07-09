// api-test.js - Simple ProWorkFlow API Tester
const https = require("https");
const http = require("http");

// Configuration - Aapke API endpoint ke hisaab se change karein
const API_CONFIG = {
  baseURL: "https://api.proworkflow.com", // Ya aapka local URL
  apiKey: "YOUR_API_KEY_HERE", // Agar API key chahiye to
  timeout: 10000, // 10 seconds timeout
};

// Colors for console output
const colors = {
  reset: "\x1b[0m",
  green: "\x1b[32m",
  red: "\x1b[31m",
  yellow: "\x1b[33m",
  blue: "\x1b[34m",
};

// Simple logger
function log(message, type = "info") {
  const prefix = {
    info: `${colors.blue}[INFO]${colors.reset}`,
    success: `${colors.green}[SUCCESS]${colors.reset}`,
    error: `${colors.red}[ERROR]${colors.reset}`,
    warning: `${colors.yellow}[WARNING]${colors.reset}`,
  };
  console.log(`${prefix[type]} ${message}`);
}

// Function to test API connection
function testAPI() {
  log("Starting ProWorkFlow API Test...", "info");
  console.log("=".repeat(50));

  // Check if API key is set
  if (API_CONFIG.apiKey === "YOUR_API_KEY_HERE") {
    log("No API key provided. Testing without authentication...", "warning");
  }

  // Test 1: Check if API is reachable
  testConnection();
}

// Test 1: Connection Test
function testConnection() {
  log("Test 1: Checking API connection...", "info");

  const url = new URL(API_CONFIG.baseURL);
  const options = {
    hostname: url.hostname,
    port: url.port || (url.protocol === "https:" ? 443 : 80),
    path: "/",
    method: "GET",
    timeout: API_CONFIG.timeout,
  };

  const protocol = url.protocol === "https:" ? https : http;
  const req = protocol.request(options, (res) => {
    console.log(`  Status Code: ${res.statusCode}`);

    if (res.statusCode >= 200 && res.statusCode < 300) {
      log("✅ API is reachable!", "success");
      testAPIResponse();
    } else if (res.statusCode === 401 || res.statusCode === 403) {
      log("⚠️ API requires authentication (401/403)", "warning");
      log("  This is normal if API key is required.", "info");
      testAPIResponse();
    } else {
      log(`❌ API returned status: ${res.statusCode}`, "error");
      testAPIResponse();
    }
  });

  req.on("error", (error) => {
    log(`❌ Connection failed: ${error.message}`, "error");
    log("  Please check:", "info");
    log("  1. API URL is correct", "info");
    log("  2. Internet connection is active", "info");
    log("  3. CORS/firewall is not blocking", "info");
    checkLocalServer();
  });

  req.on("timeout", () => {
    log("❌ Connection timeout", "error");
    req.destroy();
  });

  req.end();
}

// Test 2: API Response Test
function testAPIResponse() {
  log("Test 2: Checking API response format...", "info");

  // Simple GET request to test endpoint
  const url = new URL(API_CONFIG.baseURL + "/api/v1/test"); // Adjust endpoint as needed

  const options = {
    hostname: url.hostname,
    port: url.port || (url.protocol === "https:" ? 443 : 80),
    path: url.pathname + url.search,
    method: "GET",
    headers:
      API_CONFIG.apiKey !== "YOUR_API_KEY_HERE"
        ? {
            Authorization: `Bearer ${API_CONFIG.apiKey}`,
            "Content-Type": "application/json",
          }
        : {},
    timeout: API_CONFIG.timeout,
  };

  const protocol = url.protocol === "https:" ? https : http;
  const req = protocol.request(options, (res) => {
    let data = "";

    res.on("data", (chunk) => {
      data += chunk;
    });

    res.on("end", () => {
      try {
        if (data) {
          const json = JSON.parse(data);
          log("✅ API returned valid JSON response!", "success");
          console.log("  Response:", JSON.stringify(json, null, 2).substring(0, 200) + "...");
        } else {
          log("⚠️ API returned empty response", "warning");
        }
        finalReport(true);
      } catch (error) {
        log("⚠️ API response is not valid JSON", "warning");
        log(`  Response: ${data.substring(0, 100)}`, "info");
        finalReport(true);
      }
    });
  });

  req.on("error", (error) => {
    log(`❌ Request failed: ${error.message}`, "error");
    finalReport(false);
  });

  req.on("timeout", () => {
    log("❌ Request timeout", "error");
    req.destroy();
    finalReport(false);
  });

  req.end();
}

// Check local development server
function checkLocalServer() {
  log("Checking local development server...", "info");

  const localURLs = [
    "http://localhost:3000",
    "http://localhost:8080",
    "http://localhost:5000",
    "http://localhost:8000",
  ];

  let checkedCount = 0;

  localURLs.forEach((url) => {
    const testUrl = new URL(url);
    const options = {
      hostname: testUrl.hostname,
      port: testUrl.port || 80,
      path: "/",
      method: "GET",
      timeout: 3000,
    };

    const protocol = testUrl.protocol === "https:" ? https : http;
    const req = protocol.request(options, (res) => {
      if (res.statusCode >= 200 && res.statusCode < 300) {
        log(`✅ Local server found at ${url}`, "success");
      }
      checkedCount++;
      if (checkedCount === localURLs.length) {
        finalReport(false);
      }
    });

    req.on("error", () => {
      checkedCount++;
      if (checkedCount === localURLs.length) {
        log("No local server found on common ports", "info");
        finalReport(false);
      }
    });

    req.on("timeout", () => {
      checkedCount++;
      if (checkedCount === localURLs.length) {
        finalReport(false);
      }
      req.destroy();
    });

    req.end();
  });
}

// Final Report
function finalReport(success) {
  console.log("=".repeat(50));
  if (success) {
    log("🎉 ProWorkFlow API is working correctly!", "success");
    log("All tests passed.", "info");
    log("You can start building your app.", "info");
  } else {
    log("⚠️ Some tests failed. Please check the errors above.", "error");
    log("Common solutions:", "info");
    log("1. Make sure your API URL is correct", "info");
    log("2. Check if API key is valid", "info");
    log("3. Ensure backend server is running", "info");
    log("4. Check network/firewall settings", "info");
  }
  console.log("=".repeat(50));
}

// Run the test
if (require.main === module) {
  testAPI();
}

// Export for use in other files
module.exports = { testAPI };
