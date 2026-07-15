import React, { useEffect, useState } from "react";
import { Box, Button, Typography, Paper, CircularProgress } from "@mui/material";

const ApiTest: React.FC = () => {
  const [projects, setProjects] = useState<any[]>([]);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);

  // Direct fetch function – no service, no middleware
  const fetchProjectsDirect = async () => {
    setLoading(true);
    setError(null);

    const API_KEY = "EIOM-ZTAY-VAD8-D2Y4-PWFO9JJ-AS11627"; // hardcoded for testing
    const url = "https://api.proworkflow.com/api/v4/projects";

    console.log("🚀 Starting fetch to:", url);
    console.log("🔑 Using API Key:", API_KEY);

    try {
      const response = await fetch(url, {
        method: "GET",
        headers: {
          apikey: API_KEY,
          Accept: "application/json",
        },
        // If CORS is an issue, try adding these:
        // mode: 'cors',
        // credentials: 'omit',
      });

      console.log("📡 Response status:", response.status, response.statusText);

      // Read response text first (to log raw)
      const rawText = await response.text();
      console.log("📄 Raw response text:", rawText);

      // Try to parse JSON
      let jsonData;
      try {
        jsonData = JSON.parse(rawText);
      } catch (parseErr) {
        console.warn("⚠️ Response is not valid JSON:", parseErr);
        jsonData = rawText; // fallback
      }

      console.log("✅ Parsed response:", jsonData);

      if (!response.ok) {
        throw new Error(`HTTP ${response.status}: ${rawText || response.statusText}`);
      }

      // ProWorkflow often returns { status: "success", data: [...] }
      // So extract the data array
      const projectsArray = jsonData?.data || jsonData;
      if (Array.isArray(projectsArray)) {
        setProjects(projectsArray);
        console.log("🎯 Projects set:", projectsArray.length, "items");
      } else {
        console.warn("⚠️ Response is not an array:", projectsArray);
        setProjects([]);
      }
    } catch (err: any) {
      console.error("❌ Fetch failed:", err);
      setError(err.message || "Failed to fetch projects");
    } finally {
      setLoading(false);
    }
  };

  return (
    <Box sx={{ p: 3 }}>
      <Paper sx={{ p: 3 }}>
        <Typography variant="h6" gutterBottom>
          🧪 API Test - Direct Fetch
        </Typography>
        <Button
          variant="contained"
          onClick={fetchProjectsDirect}
          disabled={loading}
          startIcon={loading ? <CircularProgress size={20} /> : null}
        >
          {loading ? "Fetching..." : "Fetch Projects (Direct)"}
        </Button>

        {error && (
          <Typography color="error" sx={{ mt: 2 }}>
            ❌ Error: {error}
          </Typography>
        )}

        <Typography variant="subtitle2" sx={{ mt: 2 }}>
          Projects ({projects.length}):
        </Typography>
        <pre style={{ background: "#f5f5f5", padding: 10, maxHeight: 300, overflow: "auto" }}>
          {JSON.stringify(projects, null, 2)}
        </pre>

        <Typography variant="caption" color="textSecondary" sx={{ display: "block", mt: 1 }}>
          Check the browser console (F12) for detailed logs.
        </Typography>
      </Paper>
    </Box>
  );
};

export default ApiTest;
