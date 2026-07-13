// test-api.js - Node.js environment mein chalayein
// Run command: node test-api.js

const fetch = require("node-fetch");

// ===================== CONFIGURATION =====================
const CONFIG = {
  API_BASE_URL: "https://api.proworkflow.com/api/v1",
  API_KEY: "5RZL-NTMZ-IANS-9DOG-PWFMS6C-TR123459",
  EMAIL: "mubeen.dev356@gmail.com",
  PASSWORD: "MubeenNawaz@356",
};

// ===================== HELPER FUNCTIONS =====================
function getBasicAuth() {
  const credentials = `${CONFIG.EMAIL}:${CONFIG.PASSWORD}`;
  return Buffer.from(credentials).toString("base64");
}

function getHeaders(includeAuth = true) {
  const headers = {
    "X-API-Key": CONFIG.API_KEY,
    "Content-Type": "application/json",
  };

  if (includeAuth) {
    headers["Authorization"] = `Basic ${getBasicAuth()}`;
  }

  return headers;
}

// ===================== API TEST FUNCTIONS =====================

// Test 1: Get Projects
async function testGetProjects() {
  console.log("\n🚀 TEST 1: Getting Projects...");
  try {
    const response = await fetch(`${CONFIG.API_BASE_URL}/projects`, {
      method: "GET",
      headers: getHeaders(true),
    });

    const data = await response.json();

    console.log(`📊 Status: ${response.status} ${response.statusText}`);
    console.log("📦 Response:", JSON.stringify(data, null, 2));

    if (response.ok) {
      console.log("✅ SUCCESS: Projects fetched successfully!");
      return data;
    } else {
      console.log("❌ ERROR: Failed to fetch projects");
      return null;
    }
  } catch (error) {
    console.error("❌ Exception:", error.message);
    return null;
  }
}

// Test 2: Get Tasks
async function testGetTasks() {
  console.log("\n🚀 TEST 2: Getting Tasks...");
  try {
    const response = await fetch(`${CONFIG.API_BASE_URL}/tasks`, {
      method: "GET",
      headers: getHeaders(true),
    });

    const data = await response.json();

    console.log(`📊 Status: ${response.status} ${response.statusText}`);
    console.log("📦 Response:", JSON.stringify(data, null, 2));

    if (response.ok) {
      console.log("✅ SUCCESS: Tasks fetched successfully!");
      return data;
    } else {
      console.log("❌ ERROR: Failed to fetch tasks");
      return null;
    }
  } catch (error) {
    console.error("❌ Exception:", error.message);
    return null;
  }
}

// Test 3: Create Task (if you have a project ID)
async function testCreateTask(projectId) {
  console.log("\n🚀 TEST 3: Creating Task...");

  if (!projectId) {
    console.log("⚠️ SKIPPED: No projectId provided. Please provide a project ID.");
    return null;
  }

  const taskData = {
    title: "Test Task from API " + new Date().toISOString(),
    description: "This is a test task created from test script",
    projectId: projectId,
    priority: "medium",
    status: "pending",
  };

  try {
    const response = await fetch(`${CONFIG.API_BASE_URL}/tasks`, {
      method: "POST",
      headers: getHeaders(true),
      body: JSON.stringify(taskData),
    });

    const data = await response.json();

    console.log(`📊 Status: ${response.status} ${response.statusText}`);
    console.log("📦 Response:", JSON.stringify(data, null, 2));

    if (response.ok) {
      console.log("✅ SUCCESS: Task created successfully!");
      return data;
    } else {
      console.log("❌ ERROR: Failed to create task");
      return null;
    }
  } catch (error) {
    console.error("❌ Exception:", error.message);
    return null;
  }
}

// Test 4: API Key Only Test (Without Basic Auth)
async function testAPIKeyOnly() {
  console.log("\n🚀 TEST 4: Testing API Key Only (without Basic Auth)...");
  try {
    const response = await fetch(`${CONFIG.API_BASE_URL}/projects`, {
      method: "GET",
      headers: getHeaders(false), // No Basic Auth
    });

    const data = await response.json();

    console.log(`📊 Status: ${response.status} ${response.statusText}`);
    console.log("📦 Response:", JSON.stringify(data, null, 2));

    if (response.ok) {
      console.log("✅ SUCCESS: API Key only works!");
      return data;
    } else {
      console.log("❌ ERROR: API Key only does NOT work (needs Basic Auth)");
      return null;
    }
  } catch (error) {
    console.error("❌ Exception:", error.message);
    return null;
  }
}

// ===================== MAIN TEST RUNNER =====================
async function runAllTests() {
  console.log("🔄 ========== STARTING API TESTS ==========");
  console.log(`📧 Email: ${CONFIG.EMAIL}`);
  console.log(`🔑 API Key: ${CONFIG.API_KEY.substring(0, 10)}...`);
  console.log(`🔗 Base URL: ${CONFIG.API_BASE_URL}`);

  // Test 1: Get Projects
  const projects = await testGetProjects();
  console.log("📋 Projects:", projects);

  // Test 2: Get Tasks
  const tasks = await testGetTasks();
  console.log("📋 Tasks:", tasks);

  // Test 3: Create Task (if projects exist)
  if (projects && projects.length > 0) {
    const projectId = projects[0].id; // Use first project
    console.log(`📌 Using Project ID: ${projectId}`);
    const newTask = await testCreateTask(projectId);
    console.log("📋 New Task:", newTask);
  }

  // Test 4: API Key Only
  await testAPIKeyOnly();

  console.log("\n✅ ========== TESTS COMPLETED ==========");
  console.log("\n📝 Summary:");
  console.log("• If Tests 1-3 passed: Your credentials are valid!");
  console.log("• If Test 4 failed: You need Basic Auth with API Key");
  console.log("• Check the responses above for specific errors.");
}

// ===================== RUN TESTS =====================
runAllTests().catch((error) => {
  console.error("❌ Fatal Error:", error);
});
