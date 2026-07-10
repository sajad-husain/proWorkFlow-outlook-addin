// app-test.js - Complete Working Version
const https = require("https");

async function getAPIKey() {
  return new Promise((resolve, reject) => {
    // ✅ FIX 1: Username = YOUR LOGIN EMAIL
    const username = "sajjadhusain8084@gmail.com"; // 👈 YAHAN APNI EMAIL DALEIN
    const password = "!@#123Sajjad";

    // ✅ FIX 2: Account URL = SIRF ACCOUNT NAME
    const accountUrl = "NexaAIFCC2";

    // ✅ FIX 3: Proper Base64 Encoding
    const credentials = Buffer.from(`${username}:${password}`).toString("base64");

    // ✅ FIX 4: URL Encoding for special characters
    const encodedUrl = encodeURIComponent(accountUrl);

    const options = {
      hostname: "api.proworkflow.net",
      path: `/login?url=${encodedUrl}`,
      method: "GET",
      headers: {
        Authorization: `Basic ${credentials}`,
        "Content-Type": "application/json",
        "User-Agent": "Node.js-API-Test", // Some APIs require this
      },
      // ✅ FIX 5: Timeout to avoid hanging
      timeout: 10000,
    };

    console.log("⏳ Fetching API Key...");
    console.log(`📡 URL: https://api.proworkflow.net/login?url=${accountUrl}`);
    console.log(`👤 Username: ${username}`);

    const req = https.request(options, (res) => {
      let data = "";

      console.log(`📊 Status Code: ${res.statusCode}`);

      res.on("data", (chunk) => {
        data += chunk;
      });

      res.on("end", () => {
        if (res.statusCode === 200) {
          try {
            const json = JSON.parse(data);
            if (json.apikey) {
              console.log("\n✅ SUCCESS!");
              console.log(`🔑 API Key: ${json.apikey}`);
              console.log(`👤 User ID: ${json.userid}`);
              console.log(`🏢 Company: ${json.companyname}`);
              console.log(`📅 Trial: ${json.trialenddate || "N/A"}`);
              resolve(json.apikey);
            } else {
              reject(new Error("API Key not found in response"));
            }
          } catch (error) {
            reject(new Error(`Invalid JSON: ${data.substring(0, 100)}`));
          }
        } else if (res.statusCode === 401) {
          console.error("\n❌ AUTHENTICATION FAILED!");
          console.error("Possible reasons:");
          console.error('1. Username is NOT "NexaAIFCC2" - it should be your EMAIL address');
          console.error("2. Password is incorrect");
          console.error('3. Account URL "NexaAIFCC2" might be wrong');
          reject(new Error("Invalid credentials"));
        } else {
          reject(new Error(`HTTP ${res.statusCode}: ${data}`));
        }
      });
    });

    req.on("error", (error) => {
      console.error("\n❌ Network Error:", error.message);
      console.error("Possible reasons:");
      console.error("1. No internet connection");
      console.error("2. Proxy/Firewall blocking");
      console.error("3. DNS resolution failed");
      reject(error);
    });

    req.on("timeout", () => {
      req.destroy();
      reject(new Error("Request timeout after 10 seconds"));
    });

    req.end();
  });
}

// Run the function
getAPIKey()
  .then((apiKey) => {
    console.log("\n💾 Save this API key in your .env file:");
    console.log(`REACT_APP_PROWORKFLOW_API_KEY=${apiKey}`);
  })
  .catch((error) => {
    console.error("\n❌ Failed:", error.message);
    console.log("\n🔧 TROUBLESHOOTING:");
    console.log("1. Go to ProWorkFlow login page");
    console.log("2. Check your email address (username)");
    console.log("3. Reset password if needed");
    console.log("4. Contact account owner for API access");
  });
