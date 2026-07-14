import React, { useState, useEffect } from "react";
import { ThemeProvider } from "@mui/material/styles";
import CssBaseline from "@mui/material/CssBaseline";
import { theme } from "../theme/theme";
import { AppRouter } from "./Router/AppRouter";
import ErrorBoundary from "./ErrorBoundary/ErrorBoundary";
import ApiKeySetup from "./ApiKeySetup/ApiKeySetup";
import { setApiKey } from "../services/proworkflow";

const App: React.FC = () => {
  const [isApiKeySet, setIsApiKeySet] = useState<boolean>(false);
  const [loading, setLoading] = useState(true);

  useEffect(() => {
    // Check if API key exists
    const savedApiKey = localStorage.getItem("proworkflow-api-key");
    if (savedApiKey) {
      setApiKey(savedApiKey);
      setIsApiKeySet(true);
    }
    setLoading(false);
  }, []);

  const handleApiKeySuccess = () => {
    setIsApiKeySet(true);
  };

  if (loading) {
    return null;
  }

  return (
    <ThemeProvider theme={theme}>
      <CssBaseline />
      <ErrorBoundary>
        {isApiKeySet ? <AppRouter /> : <ApiKeySetup open={true} onSuccess={handleApiKeySuccess} />}
      </ErrorBoundary>
    </ThemeProvider>
  );
};

export default App;
