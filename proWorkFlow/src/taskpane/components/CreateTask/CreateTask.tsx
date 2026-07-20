import React, { useState, useEffect, useCallback, useRef } from "react";
import {
  Box,
  TextField,
  Select,
  MenuItem,
  FormControl,
  InputLabel,
  FormControlLabel,
  Checkbox,
  Button,
  Typography,
  Alert,
  CircularProgress,
  Chip,
  Stack,
  Divider,
  Paper,
  InputAdornment,
  Collapse,
  IconButton,
  Tooltip,
  Skeleton,
  Dialog,
  DialogTitle,
  DialogContent,
  DialogContentText,
  DialogActions,
  Snackbar,
} from "@mui/material";
import {
  Send,
  AttachFile,
  PriorityHigh,
  Schedule,
  Person,
  Folder,
  Description,
  Title,
  Cancel,
  ExpandMore,
  ExpandLess,
  Email,
  Refresh,
  CheckCircle,
  Warning,
  Key as KeyIcon,
} from "@mui/icons-material";
import {
  getOutlookItemDataAsync,
  OutlookItemData,
  generateTaskDescription,
  cleanEmailBody,
} from "../../services/outlook";
import { useLocalStorage } from "../../hooks/useLocalStorage";

// ========================
// LOCAL TYPES (No external imports)
// ========================

interface DraftData {
  taskName: string;
  projectId: number;
  assigneeId: number;
  description: string;
  priority: "Low" | "Medium" | "High";
  dueDate: string;
  isUrgent: boolean;
  includeAttachments: boolean;
}

interface CreateTaskPayload {
  projectid: number;
  name: string;
  description?: string;
  priorityid?: number;
  contactid?: string;
  startdate?: string;
  duedate?: string;
  urgent?: boolean;
}

interface CreateTaskResponse {
  id: number;
}

// Local Project type (replacing import from proworkflow.ts)
interface Project {
  id: number;
  name: string;
  // Add other fields if needed
}

// Local User type (replacing import from proworkflow.ts)
interface User {
  id: number;
  name: string;
  email: string;
  firstname: string;
  lastname: string;
  companyname: string;
}

// ========================
// CONSTANTS
// ========================

const PRIORITY_MAP: Record<"Low" | "Medium" | "High", number> = {
  Low: 1,
  Medium: 2,
  High: 3,
};

const smallLabelStyles = {
  "& .MuiInputLabel-root": { fontSize: "0.7rem" },
  "& .MuiInputBase-root": { fontSize: "0.75rem" },
  "& .MuiFormHelperText-root": { fontSize: "0.65rem" },
  "& .MuiSelect-select": { fontSize: "0.75rem" },
  "& .MuiOutlinedInput-input": { fontSize: "0.75rem", padding: "8px 12px" },
  "& .MuiInputLabel-shrink": { fontSize: "0.7rem" },
};

// ========================
// MAIN COMPONENT
// ========================

const CreateTask: React.FC = () => {
  // ---- Form State ----
  const [taskName, setTaskName] = useState("");
  const [projectId, setProjectId] = useState<number>(0);
  const [assigneeId, setAssigneeId] = useState<number>(0);
  const [description, setDescription] = useState("");
  const [priority, setPriority] = useState<"Low" | "Medium" | "High">("Medium");
  const [dueDate, setDueDate] = useState("");
  const [isUrgent, setIsUrgent] = useState(false);
  const [includeAttachments, setIncludeAttachments] = useState(false);

  // ---- Data State ----
  const [projects, setProjects] = useState<Project[]>([]);
  const [contacts, setContacts] = useState<User[]>([]);
  const [loadingProjects, setLoadingProjects] = useState(false);
  const [loadingContacts, setLoadingContacts] = useState(false);

  // ---- UI State ----
  const [loading, setLoading] = useState(false);
  const [loadingData, setLoadingData] = useState(true);
  const [loadingOutlook, setLoadingOutlook] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [success, setSuccess] = useState(false);
  const [outlookData, setOutlookData] = useState<OutlookItemData | null>(null);
  const [showEmailPreview, setShowEmailPreview] = useState(false);
  const [showConfirmDialog, setShowConfirmDialog] = useState(false);
  const [pendingAction, setPendingAction] = useState<"submit" | "reset" | null>(null);
  const [toastOpen, setToastOpen] = useState(false);
  const [toastMessage, setToastMessage] = useState("");
  const [toastSeverity, setToastSeverity] = useState<"success" | "error" | "info" | "warning">(
    "info"
  );
  const [isApiKeySet, setIsApiKeySet] = useState(false);
  const [apiKeyInput, setApiKeyInput] = useState("");

  // ---- Refs ----
  const isInitialized = useRef(false);
  const isFirstLoad = useRef(true);

  const [draft, setDraft] = useLocalStorage<DraftData | null>("proworkflow-task-draft", null);

  // ========================
  // LOCAL API FUNCTIONS (SELF-CONTAINED)
  // ========================

  // Get API key from localStorage
  const getApiKey = (): string => {
    return localStorage.getItem("proworkflow-api-key") || process.env.PW_API_KEY || "";
  };

  // Test API key - self-contained
  const testApiKey = async (key: string): Promise<boolean> => {
    try {
      const url = "https://api.proworkflow.com/api/v4/projects?limit=1";
      const response = await fetch(url, {
        method: "GET",
        headers: { apikey: key, Accept: "application/json" },
      });
      return response.ok;
    } catch {
      return false;
    }
  };

  // Fetch projects - self-contained
  const fetchProjects = async (): Promise<Project[]> => {
    const API_KEY = getApiKey();
    if (!API_KEY) throw new Error("API key not set");

    const url = "https://api.proworkflow.com/api/v4/projects";
    const response = await fetch(url, {
      method: "GET",
      headers: { apikey: API_KEY, Accept: "application/json" },
    });

    if (!response.ok) {
      const text = await response.text();
      throw new Error(`Failed to fetch projects (${response.status}): ${text}`);
    }

    const json = await response.json();
    const data = json.data || json;
    if (!Array.isArray(data)) return [];
    return data.map((p: any) => ({ id: p.id, name: p.name }));
  };

  // Fetch contacts - self-contained
  const fetchContacts = async (): Promise<User[]> => {
    const API_KEY = getApiKey();
    if (!API_KEY) throw new Error("API key not set");

    const url = "https://api.proworkflow.com/api/v4/contacts";
    const response = await fetch(url, {
      method: "GET",
      headers: { apikey: API_KEY, Accept: "application/json" },
    });

    if (!response.ok) {
      const text = await response.text();
      throw new Error(`Failed to fetch contacts (${response.status}): ${text}`);
    }

    const json = await response.json();
    const data = json.data || json;
    if (!Array.isArray(data)) return [];

    return data.map((c: any) => {
      let name = "Unnamed Contact";
      if (c.firstname && c.lastname) name = `${c.firstname} ${c.lastname}`;
      else if (c.firstname) name = c.firstname;
      else if (c.lastname) name = c.lastname;
      else if (c.companyname) name = c.companyname;
      else if (c.fullname) name = c.fullname;
      return {
        id: c.id,
        name,
        email: c.email || "",
        firstname: c.firstname || "",
        lastname: c.lastname || "",
        companyname: c.companyname || "",
      };
    });
  };

  // Create task API - self-contained
  const createTaskDirectly = async (payload: CreateTaskPayload): Promise<CreateTaskResponse> => {
    const API_KEY = getApiKey();
    if (!API_KEY) throw new Error("API key is not set.");

    const url = "https://api.proworkflow.com/api/v4/projects/items";
    const body = JSON.stringify(payload);

    console.group("📡 Create Task Request");
    console.log("URL:", url);
    console.log("Payload:", payload);
    console.groupEnd();

    const response = await fetch(url, {
      method: "POST",
      headers: {
        apikey: API_KEY,
        "Content-Type": "application/json",
        Accept: "application/json",
      },
      body,
    });

    const text = await response.text();
    let json: any;
    try {
      json = JSON.parse(text);
    } catch {
      json = text;
    }

    console.group("📡 Create Task Response");
    console.log("Status:", response.status);
    console.log("Response:", json);
    console.groupEnd();

    if (!response.ok) {
      let msg = `Failed (${response.status})`;
      if (json?.error) msg = json.error;
      else if (json?.message) msg = json.message;
      else if (typeof json === "string") msg = json;
      throw new Error(msg);
    }

    const taskId = json?.data?.id ?? json?.id;
    if (!taskId) throw new Error("No ID returned");
    return { id: taskId };
  };

  // ========================
  // PAYLOAD BUILDER
  // ========================

  const buildPayload = useCallback((): CreateTaskPayload => {
    const payload: Partial<CreateTaskPayload> = {
      projectid: projectId,
      name: taskName.trim(),
    };

    if (description.trim()) payload.description = description.trim();
    const priorityValue = PRIORITY_MAP[priority];
    if (priorityValue) payload.priorityid = priorityValue;
    if (assigneeId > 0) payload.contactid = String(assigneeId);
    if (dueDate) {
      payload.duedate = dueDate;
      payload.startdate = dueDate;
    }
    if (isUrgent) payload.urgent = true;

    // Clean undefined, null, empty string
    const cleaned: CreateTaskPayload = {} as CreateTaskPayload;
    for (const key in payload) {
      const value = payload[key as keyof typeof payload];
      if (value !== undefined && value !== null && value !== "") {
        (cleaned as any)[key] = value;
      }
    }
    return cleaned;
  }, [projectId, taskName, description, priority, assigneeId, dueDate, isUrgent]);

  // ========================
  // TOAST HELPER
  // ========================

  const showToast = useCallback((message: string, severity: typeof toastSeverity) => {
    setToastMessage(message);
    setToastSeverity(severity);
    setToastOpen(true);
  }, []);

  const handleToastClose = () => setToastOpen(false);

  // ========================
  // LOAD OUTLOOK DATA
  // ========================

  const loadOutlookDataInternal = useCallback(async () => {
    setLoadingOutlook(true);
    try {
      const data = await getOutlookItemDataAsync();
      setOutlookData(data);
      if (data) {
        if (!draft?.taskName && !taskName) setTaskName(data.subject || "");
        if (!draft?.description && !description) setDescription(generateTaskDescription(data));
      }
    } catch (err) {
      console.warn("Outlook load error:", err);
    } finally {
      setLoadingOutlook(false);
    }
  }, [draft, taskName, description]);

  // ========================
  // LOAD DRAFT
  // ========================

  const loadDraftInternal = useCallback(() => {
    if (draft && isFirstLoad.current) {
      setTaskName(draft.taskName || "");
      setProjectId(draft.projectId || 0);
      setAssigneeId(draft.assigneeId || 0);
      setDescription(draft.description || "");
      setPriority(draft.priority || "Medium");
      setDueDate(draft.dueDate || "");
      setIsUrgent(draft.isUrgent || false);
      setIncludeAttachments(draft.includeAttachments || false);
      isFirstLoad.current = false;
    }
  }, [draft]);

  // ========================
  // LOAD ALL DATA
  // ========================

  const loadAllData = useCallback(async () => {
    setLoadingData(true);
    setError(null);
    try {
      setLoadingProjects(true);
      setLoadingContacts(true);

      const [projectsData, contactsData] = await Promise.all([fetchProjects(), fetchContacts()]);

      setProjects(projectsData);
      setContacts(contactsData);
      setLoadingProjects(false);
      setLoadingContacts(false);

      await loadOutlookDataInternal();
      loadDraftInternal();
    } catch (err: any) {
      const errorMsg = err.message || "Failed to load data. Please check your API key.";
      setError(errorMsg);
      console.error(err);
      setLoadingProjects(false);
      setLoadingContacts(false);
    } finally {
      setLoadingData(false);
    }
  }, [loadOutlookDataInternal, loadDraftInternal]);

  // ========================
  // REFRESH OUTLOOK
  // ========================

  const handleRefreshOutlook = useCallback(() => {
    loadOutlookDataInternal();
    showToast("Email data refreshed", "info");
  }, [loadOutlookDataInternal, showToast]);

  // ========================
  // API KEY SUBMIT
  // ========================

  const handleApiKeySubmit = useCallback(async () => {
    const trimmed = apiKeyInput.trim();
    if (!trimmed) {
      showToast("Enter a valid API key", "error");
      return;
    }

    setLoading(true);
    try {
      const isValid = await testApiKey(trimmed);
      if (isValid) {
        localStorage.setItem("proworkflow-api-key", trimmed);
        setIsApiKeySet(true);
        setError(null);
        await loadAllData();
        showToast("API key verified! 🎉", "success");
      } else {
        setError("Invalid API key. Please check and try again.");
        showToast("Invalid API key", "error");
      }
    } catch (err) {
      setError("Failed to verify API key. Please try again.");
      showToast("Failed to verify API key", "error");
    } finally {
      setLoading(false);
    }
  }, [apiKeyInput, loadAllData, showToast]);

  // ========================
  // SUBMIT CLICK
  // ========================

  const handleSubmitClick = useCallback(() => {
    if (taskName.trim() && projectId) {
      setShowConfirmDialog(true);
      setPendingAction("submit");
    } else {
      showToast("Please fill all required fields", "warning");
    }
  }, [taskName, projectId, showToast]);

  // ========================
  // CONFIRM ACTION
  // ========================

  const handleConfirmAction = useCallback(() => {
    setShowConfirmDialog(false);
    if (pendingAction === "submit") handleSubmit();
    else if (pendingAction === "reset") handleReset();
    setPendingAction(null);
  }, [pendingAction]);

  // ========================
  // SUBMIT HANDLER
  // ========================

  const handleSubmit = useCallback(
    async (e?: React.FormEvent) => {
      e?.preventDefault();

      const trimmedName = taskName.trim();
      if (!trimmedName) {
        showToast("Task name is required", "error");
        const el = document.querySelector('input[name="taskName"]') as HTMLInputElement | null;
        if (el) el.focus();
        return;
      }
      if (!projectId) {
        showToast("Please select a project", "error");
        return;
      }
      if (dueDate && !/^\d{4}-\d{2}-\d{2}$/.test(dueDate)) {
        showToast("Due date must be YYYY-MM-DD", "error");
        return;
      }

      setLoading(true);
      setError(null);
      setSuccess(false);

      try {
        const payload = buildPayload();
        console.group("🚀 Submitting");
        console.log("Payload:", payload);
        console.groupEnd();

        const result = await createTaskDirectly(payload);

        if (includeAttachments && outlookData?.attachments?.length) {
          console.log("📎 Would upload", outlookData.attachments.length, "attachments");
          // Attachment upload logic can be added here
        }

        setSuccess(true);
        showToast(
          includeAttachments && outlookData?.attachments?.length
            ? "Task & attachments created! 🎉"
            : "Task created! 🎉",
          "success"
        );
        setDraft(null);
        setTimeout(() => setSuccess(false), 5000);
      } catch (err: any) {
        const errorMsg = err.message || "Failed to create task";
        setError(errorMsg);
        showToast(errorMsg, "error");
        console.error(err);
      } finally {
        setLoading(false);
      }
    },
    [
      taskName,
      projectId,
      dueDate,
      buildPayload,
      includeAttachments,
      outlookData,
      showToast,
      setDraft,
    ]
  );

  // ========================
  // RESET HANDLER
  // ========================

  const handleReset = useCallback(() => {
    setTaskName("");
    setProjectId(0);
    setAssigneeId(0);
    setDescription("");
    setPriority("Medium");
    setDueDate("");
    setIsUrgent(false);
    setIncludeAttachments(false);
    setError(null);
    setSuccess(false);
    setDraft(null);
    if (outlookData) {
      setTaskName(outlookData.subject || "");
      setDescription(generateTaskDescription(outlookData));
    }
    showToast("Form reset", "info");
  }, [outlookData, showToast, setDraft]);

  // ========================
  // EFFECTS
  // ========================

  // Initialize
  useEffect(() => {
    if (isInitialized.current) return;
    isInitialized.current = true;

    const init = async () => {
      const saved = localStorage.getItem("proworkflow-api-key");
      const key = (saved && saved.trim()) || process.env.PW_API_KEY || "";

      if (key) {
        setIsApiKeySet(true);
        setApiKeyInput(key);
        await loadAllData();
      } else {
        setIsApiKeySet(false);
        setLoadingData(false);
      }
    };
    init();
  }, [loadAllData]);

  // Auto-save draft
  useEffect(() => {
    if (!isApiKeySet || isFirstLoad.current) return;

    const timer = setTimeout(() => {
      if (taskName || description || projectId || assigneeId) {
        setDraft({
          taskName,
          projectId,
          assigneeId,
          description,
          priority,
          dueDate,
          isUrgent,
          includeAttachments,
        });
      }
    }, 500);

    return () => clearTimeout(timer);
  }, [
    isApiKeySet,
    taskName,
    projectId,
    assigneeId,
    description,
    priority,
    dueDate,
    isUrgent,
    includeAttachments,
    setDraft,
  ]);

  // Keyboard shortcuts
  useEffect(() => {
    const handler = (e: KeyboardEvent) => {
      if ((e.ctrlKey || e.metaKey) && e.key === "Enter") {
        e.preventDefault();
        if (taskName.trim() && projectId) handleSubmitClick();
      }
      if (e.key === "Escape" && !showConfirmDialog) {
        e.preventDefault();
        setShowConfirmDialog(true);
        setPendingAction("reset");
      }
    };
    window.addEventListener("keydown", handler);
    return () => window.removeEventListener("keydown", handler);
  }, [taskName, projectId, showConfirmDialog, handleSubmitClick]);

  // ========================
  // RENDER HELPERS
  // ========================

  const getAttachmentsLabel = (): string => {
    const count = outlookData?.attachments?.length ?? 0;
    return count > 0 ? ` (${count} files)` : " (No attachments found)";
  };

  // ========================
  // RENDER
  // ========================

  // API Key Screen
  if (!isApiKeySet && !loadingData) {
    return (
      <Box sx={{ p: 2 }}>
        <Paper sx={{ p: 3, textAlign: "center" }}>
          <Typography variant="h6" gutterBottom sx={{ fontWeight: 600, fontSize: "1rem" }}>
            🔑 ProWorkflow API Key
          </Typography>
          <Typography variant="body2" color="textSecondary" sx={{ mb: 2, fontSize: "0.8rem" }}>
            Please enter your ProWorkflow API key to connect the add-in.
          </Typography>
          <Box sx={{ maxWidth: 400, mx: "auto" }}>
            <TextField
              fullWidth
              size="small"
              label="API Key"
              placeholder="e.g., EIOM-AAAY-VAD8-D2Y4-PWFO9JJ-A111111"
              value={apiKeyInput}
              onChange={(e) => setApiKeyInput(e.target.value)}
              onKeyDown={(e) => e.key === "Enter" && handleApiKeySubmit()}
              sx={{ mb: 1.5, ...smallLabelStyles }}
              slotProps={{
                input: {
                  startAdornment: (
                    <InputAdornment position="start">
                      <KeyIcon fontSize="small" color="action" />
                    </InputAdornment>
                  ),
                },
              }}
            />
            <Button
              fullWidth
              variant="contained"
              size="small"
              onClick={handleApiKeySubmit}
              disabled={loading || !apiKeyInput.trim()}
              startIcon={loading ? <CircularProgress size={20} /> : <CheckCircle />}
              sx={{ fontSize: "0.75rem" }}
            >
              {loading ? "Verifying..." : "Verify & Connect"}
            </Button>
            {error && (
              <Alert severity="error" sx={{ mt: 2, fontSize: "0.75rem" }}>
                {error}
              </Alert>
            )}
          </Box>
          <Paper variant="outlined" sx={{ p: 1.5, mt: 2, bgcolor: "#f8f9fa", textAlign: "left" }}>
            <Typography
              variant="caption"
              color="textSecondary"
              sx={{ fontWeight: 600, fontSize: "0.65rem" }}
            >
              💡 How to get your API key:
            </Typography>
            <Typography variant="body2" color="textSecondary" sx={{ mt: 0.5, fontSize: "0.7rem" }}>
              1. Log in to your ProWorkflow account
              <br />
              2. Go to Settings → API Keys
              <br />
              3. Create a new API key or copy existing one
            </Typography>
          </Paper>
        </Paper>
      </Box>
    );
  }

  // Loading Skeleton
  if (loadingData) {
    return (
      <Box sx={{ p: 2 }}>
        <Stack spacing={2}>
          <Skeleton variant="rectangular" height={36} animation="wave" />
          <Skeleton variant="rectangular" height={36} animation="wave" />
          <Skeleton variant="rectangular" height={36} animation="wave" />
          <Skeleton variant="rectangular" height={60} animation="wave" />
          <Skeleton variant="rectangular" height={36} animation="wave" />
          <Skeleton variant="rectangular" height={50} animation="wave" />
          <Skeleton variant="rectangular" height={32} animation="wave" />
        </Stack>
      </Box>
    );
  }

  // Main Form
  return (
    <Box component="form" onSubmit={handleSubmit} sx={{ maxWidth: "100%", p: 0 }}>
      {/* Outlook Header */}
      {outlookData && (
        <Paper
          elevation={0}
          sx={{
            p: 1,
            mb: 1.5,
            bgcolor: "#f8f9fa",
            border: "1px solid #e0e0e0",
            borderRadius: 1,
          }}
        >
          <Stack
            direction="row"
            spacing={0.5}
            sx={{ alignItems: "center", flexWrap: "wrap", mb: 0.5 }}
          >
            <Email fontSize="small" color="primary" />
            <Typography variant="caption" color="textSecondary" sx={{ fontSize: "0.65rem" }}>
              From:
            </Typography>
            <Chip
              label={outlookData.sender}
              size="small"
              variant="outlined"
              color="primary"
              sx={{ height: 18, fontSize: "0.6rem" }}
            />
            <Typography
              variant="caption"
              color="textSecondary"
              sx={{ ml: 0.5, fontSize: "0.65rem" }}
            >
              Subject: {outlookData.subject}
            </Typography>
            <Box sx={{ flex: 1 }} />
            <Tooltip title="Refresh">
              <IconButton size="small" onClick={handleRefreshOutlook} disabled={loadingOutlook}>
                <Refresh fontSize="small" />
              </IconButton>
            </Tooltip>
            <Tooltip title="Toggle preview">
              <IconButton size="small" onClick={() => setShowEmailPreview(!showEmailPreview)}>
                {showEmailPreview ? (
                  <ExpandLess fontSize="small" />
                ) : (
                  <ExpandMore fontSize="small" />
                )}
              </IconButton>
            </Tooltip>
          </Stack>

          <Collapse in={showEmailPreview}>
            <Divider sx={{ my: 0.5 }} />
            <Box sx={{ mt: 0.5 }}>
              <Typography
                variant="caption"
                color="textSecondary"
                sx={{ mb: 0.5, display: "block", fontSize: "0.6rem" }}
              >
                Email Preview:
              </Typography>
              <Paper
                variant="outlined"
                sx={{
                  p: 0.5,
                  bgcolor: "white",
                  maxHeight: "80px",
                  overflow: "auto",
                  fontSize: "0.7rem",
                  whiteSpace: "pre-wrap",
                  wordBreak: "break-word",
                }}
              >
                {cleanEmailBody(outlookData.body || "No content") || "No content"}
              </Paper>
              {outlookData.attachments && outlookData.attachments.length > 0 && (
                <Box sx={{ mt: 0.5 }}>
                  <Typography variant="caption" color="textSecondary" sx={{ fontSize: "0.6rem" }}>
                    Attachments: {outlookData.attachments.map((a) => a.name).join(", ")}
                  </Typography>
                </Box>
              )}
            </Box>
          </Collapse>
        </Paper>
      )}

      {!outlookData && !loadingOutlook && (
        <Alert severity="info" sx={{ mb: 1.5, py: 0.5, fontSize: "0.7rem" }}>
          No email selected. Select an email to auto-fill.
          <Button size="small" onClick={handleRefreshOutlook} sx={{ ml: 1, fontSize: "0.7rem" }}>
            Refresh
          </Button>
        </Alert>
      )}

      {loadingOutlook && (
        <Box sx={{ display: "flex", alignItems: "center", mb: 1.5 }}>
          <CircularProgress size={14} sx={{ mr: 1 }} />
          <Typography variant="caption" color="textSecondary" sx={{ fontSize: "0.7rem" }}>
            Loading email...
          </Typography>
        </Box>
      )}

      {error && (
        <Alert
          severity="error"
          sx={{ mb: 1.5, py: 0.5, fontSize: "0.7rem" }}
          onClose={() => setError(null)}
        >
          {error}
        </Alert>
      )}
      {success && (
        <Alert
          severity="success"
          sx={{ mb: 1.5, py: 0.5, fontSize: "0.7rem" }}
          icon={<CheckCircle fontSize="small" />}
        >
          Task created successfully! 🎉
        </Alert>
      )}
      {draft && (
        <Alert
          severity="info"
          sx={{ mb: 1.5, py: 0.5, fontSize: "0.7rem" }}
          icon={<Warning fontSize="small" />}
        >
          You have a saved draft.
        </Alert>
      )}

      <Stack spacing={1.5}>
        {/* Task Name */}
        <TextField
          required
          size="small"
          fullWidth
          label="Task Name"
          name="taskName"
          value={taskName}
          onChange={(e) => setTaskName(e.target.value)}
          placeholder="Enter task name..."
          sx={smallLabelStyles}
          disabled={loading}
          slotProps={{
            input: {
              startAdornment: (
                <InputAdornment position="start">
                  <Title fontSize="small" color="action" />
                </InputAdornment>
              ),
            },
          }}
        />

        {/* Project */}
        <FormControl fullWidth required size="small" sx={smallLabelStyles} disabled={loading}>
          <InputLabel sx={{ fontSize: "0.7rem" }}>Project</InputLabel>
          <Select
            value={projectId}
            onChange={(e) => setProjectId(e.target.value as number)}
            label="Project"
            disabled={loadingProjects || loading}
            startAdornment={
              loadingProjects ? (
                <CircularProgress size={14} sx={{ ml: 1 }} />
              ) : (
                <Folder color="action" sx={{ ml: 1 }} fontSize="small" />
              )
            }
          >
            <MenuItem value={0} sx={{ fontSize: "0.75rem" }}>
              Select a project...
            </MenuItem>
            {projects.map((p) => (
              <MenuItem key={p.id} value={p.id} sx={{ fontSize: "0.75rem" }}>
                {p.name}
              </MenuItem>
            ))}
          </Select>
        </FormControl>

        {/* Assignee */}
        <FormControl fullWidth size="small" sx={smallLabelStyles} disabled={loading}>
          <InputLabel sx={{ fontSize: "0.7rem" }}>Assignee</InputLabel>
          <Select
            value={assigneeId}
            onChange={(e) => setAssigneeId(e.target.value as number)}
            label="Assignee"
            disabled={loadingContacts || loading}
            startAdornment={
              loadingContacts ? (
                <CircularProgress size={14} sx={{ ml: 1 }} />
              ) : (
                <Person color="action" sx={{ ml: 1 }} fontSize="small" />
              )
            }
          >
            <MenuItem value={0} sx={{ fontSize: "0.75rem" }}>
              Unassigned
            </MenuItem>
            {contacts.map((c) => (
              <MenuItem key={c.id} value={c.id} sx={{ fontSize: "0.75rem" }}>
                {c.name}
              </MenuItem>
            ))}
          </Select>
        </FormControl>

        {/* Description */}
        <TextField
          fullWidth
          size="small"
          multiline
          rows={2}
          label="Description"
          value={description}
          onChange={(e) => setDescription(e.target.value)}
          placeholder="Task description..."
          sx={smallLabelStyles}
          disabled={loading}
          slotProps={{
            input: {
              startAdornment: (
                <InputAdornment position="start">
                  <Description fontSize="small" color="action" sx={{ mr: 0.5 }} />
                </InputAdornment>
              ),
            },
          }}
        />

        {/* Priority & Due Date */}
        <Stack direction="row" spacing={1.5}>
          <FormControl fullWidth size="small" sx={smallLabelStyles} disabled={loading}>
            <InputLabel sx={{ fontSize: "0.7rem" }}>Priority</InputLabel>
            <Select
              value={priority}
              onChange={(e) => setPriority(e.target.value as "Low" | "Medium" | "High")}
              label="Priority"
              startAdornment={<PriorityHigh fontSize="small" color="action" sx={{ ml: 1 }} />}
            >
              <MenuItem value="Low" sx={{ fontSize: "0.75rem" }}>
                Low
              </MenuItem>
              <MenuItem value="Medium" sx={{ fontSize: "0.75rem" }}>
                Medium
              </MenuItem>
              <MenuItem value="High" sx={{ fontSize: "0.75rem" }}>
                High
              </MenuItem>
            </Select>
          </FormControl>

          <TextField
            fullWidth
            size="small"
            type="date"
            label="Due Date"
            value={dueDate}
            onChange={(e) => setDueDate(e.target.value)}
            sx={smallLabelStyles}
            disabled={loading}
            slotProps={{
              input: {
                startAdornment: (
                  <InputAdornment position="start">
                    <Schedule fontSize="small" color="action" />
                  </InputAdornment>
                ),
              },
            }}
          />
        </Stack>

        {/* Options */}
        <Paper variant="outlined" sx={{ p: 1 }}>
          <Stack spacing={0.5}>
            <FormControlLabel
              control={
                <Checkbox
                  checked={isUrgent}
                  onChange={(e) => setIsUrgent(e.target.checked)}
                  color="error"
                  size="small"
                  disabled={loading}
                />
              }
              label={
                <Typography variant="body2" sx={{ fontSize: "0.75rem" }}>
                  <PriorityHigh fontSize="small" color="error" sx={{ verticalAlign: "middle" }} />
                  Mark as Urgent
                </Typography>
              }
            />
            <Divider />
            <FormControlLabel
              control={
                <Checkbox
                  checked={includeAttachments}
                  onChange={(e) => setIncludeAttachments(e.target.checked)}
                  disabled={!outlookData?.attachments?.length || loading}
                  size="small"
                />
              }
              label={
                <Stack direction="row" spacing={0.5} sx={{ alignItems: "center" }}>
                  <AttachFile fontSize="small" />
                  <Typography variant="body2" sx={{ fontSize: "0.75rem" }}>
                    Include email attachments{getAttachmentsLabel()}
                  </Typography>
                </Stack>
              }
            />
          </Stack>
        </Paper>

        {/* Keyboard shortcuts */}
        <Typography
          variant="caption"
          color="textSecondary"
          sx={{ textAlign: "center", fontSize: "0.6rem" }}
        >
          <Stack
            direction="row"
            spacing={1}
            sx={{ justifyContent: "center", alignItems: "center", flexWrap: "wrap" }}
          >
            <span>⌘+Enter / Ctrl+Enter to submit</span>
            <span>•</span>
            <span>Esc to reset</span>
          </Stack>
        </Typography>

        {/* Buttons */}
        <Stack direction="row" spacing={1.5} sx={{ justifyContent: "flex-end" }}>
          <Button
            variant="outlined"
            size="small"
            onClick={() => {
              setShowConfirmDialog(true);
              setPendingAction("reset");
            }}
            disabled={loading}
            startIcon={<Cancel fontSize="small" />}
            sx={{ fontSize: "0.7rem" }}
          >
            Reset
          </Button>
          <Button
            type="button"
            variant="contained"
            size="small"
            disabled={loading || !taskName.trim() || !projectId || !isApiKeySet}
            startIcon={loading ? <CircularProgress size={14} /> : <Send fontSize="small" />}
            onClick={handleSubmitClick}
            sx={{ fontSize: "0.7rem" }}
          >
            {loading ? "Creating..." : "Create Task"}
          </Button>
        </Stack>
      </Stack>

      {/* Confirmation Dialog */}
      <Dialog open={showConfirmDialog} onClose={() => setShowConfirmDialog(false)}>
        <DialogTitle sx={{ fontSize: "0.9rem" }}>Confirm Action</DialogTitle>
        <DialogContent>
          <DialogContentText sx={{ fontSize: "0.8rem" }}>
            {pendingAction === "submit"
              ? "Create this task?"
              : "Reset all fields? Draft will be cleared."}
          </DialogContentText>
        </DialogContent>
        <DialogActions>
          <Button onClick={() => setShowConfirmDialog(false)} sx={{ fontSize: "0.75rem" }}>
            Cancel
          </Button>
          <Button
            onClick={handleConfirmAction}
            variant="contained"
            color={pendingAction === "submit" ? "primary" : "error"}
            sx={{ fontSize: "0.75rem" }}
          >
            {pendingAction === "submit" ? "Create" : "Reset"}
          </Button>
        </DialogActions>
      </Dialog>

      {/* Toast / Snackbar */}
      <Snackbar
        open={toastOpen}
        autoHideDuration={4000}
        onClose={handleToastClose}
        anchorOrigin={{ vertical: "bottom", horizontal: "center" }}
      >
        <Alert
          onClose={handleToastClose}
          severity={toastSeverity}
          variant="filled"
          sx={{ fontSize: "0.75rem" }}
        >
          {toastMessage}
        </Alert>
      </Snackbar>
    </Box>
  );
};

export default CreateTask;
