import React, { useState } from "react";
import {
  Box,
  Dialog,
  DialogTitle,
  DialogContent,
  DialogContentText,
  TextField,
  Button,
  Stack,
  Alert,
  CircularProgress,
  Typography,
  Paper,
} from "@mui/material";
import { Key, CheckCircle, Error } from "@mui/icons-material";
import { proWorkflowApi, setApiKey } from "../../services/proworkflow";

interface ApiKeySetupProps {
  open: boolean;
  onSuccess: () => void;
}

const ApiKeySetup: React.FC<ApiKeySetupProps> = ({ open, onSuccess }) => {
  const [apiKey, setApiKeyState] = useState("");
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [success, setSuccess] = useState(false);

  const handleVerify = async () => {
    if (!apiKey.trim()) {
      setError("Please enter your API key");
      return;
    }

    setLoading(true);
    setError(null);
    setSuccess(false);

    try {
      const isValid = await proWorkflowApi.testApiKey(apiKey.trim());

      if (isValid) {
        // Save API key
        localStorage.setItem("proworkflow-api-key", apiKey.trim());
        setApiKey(apiKey.trim());
        setSuccess(true);
        setTimeout(() => {
          onSuccess();
        }, 1000);
      } else {
        setError("Invalid API key. Please check and try again.");
      }
    } catch (err) {
      setError("Failed to verify API key. Please try again.");
      console.error(err);
    } finally {
      setLoading(false);
    }
  };

  return (
    <Dialog open={open} maxWidth="sm" fullWidth>
      <DialogTitle>
        <Stack direction="row" spacing={1} sx={{ alignItems: "center" }}>
          <Key color="primary" />
          <Typography variant="h6">ProWorkflow API Key</Typography>
        </Stack>
      </DialogTitle>
      <DialogContent>
        <DialogContentText sx={{ mb: 3 }}>
          Please enter your ProWorkflow API key to connect the add-in. You can find your API key in
          your ProWorkflow account settings.
        </DialogContentText>

        <Paper variant="outlined" sx={{ p: 2, bgcolor: "#f8f9fa", mb: 3 }}>
          <Typography variant="caption" color="textSecondary">
            💡 How to get your API key:
          </Typography>
          <Typography variant="body2" color="textSecondary" sx={{ mt: 1 }}>
            1. Log in to your ProWorkflow account
            <br />
            2. Go to Settings → API Keys
            <br />
            3. Create a new API key or copy existing one
          </Typography>
        </Paper>

        <TextField
          fullWidth
          label="API Key"
          value={apiKey}
          onChange={(e) => setApiKeyState(e.target.value)}
          placeholder="Enter your API key (e.g., EIOM-ZTAY-VAD8-D2Y4-PWFO9JJ-AS11627)"
          disabled={loading || success}
          sx={{ mb: 2 }}
          slotProps={{
            input: {
              startAdornment: <Key color="action" sx={{ mr: 1 }} />,
            },
          }}
        />

        {error && (
          <Alert severity="error" sx={{ mb: 2 }}>
            {error}
          </Alert>
        )}

        {success && (
          <Alert severity="success" sx={{ mb: 2 }} icon={<CheckCircle />}>
            API key verified successfully! 🎉
          </Alert>
        )}

        <Button
          fullWidth
          variant="contained"
          onClick={handleVerify}
          disabled={loading || success || !apiKey.trim()}
          startIcon={loading ? <CircularProgress size={20} /> : <CheckCircle />}
          sx={{ mt: 1 }}
        >
          {loading ? "Verifying..." : "Verify & Connect"}
        </Button>
      </DialogContent>
    </Dialog>
  );
};

export default ApiKeySetup;
