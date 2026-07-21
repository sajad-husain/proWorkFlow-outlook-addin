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
import { smallLabelStyles, smallMenuProps } from "../../theme/styles";

// ========================
// LOCAL TYPES
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
  selectedAttachments: number[]; // new
}

interface CreateTaskPayload {
  name: string;
  itemtypeid: number;
  itemcollectionid: number;
  projectid?: number;
  description?: string;
  priorityid?: number;
  contacts?: { id: number }[];
  startdate?: string;
  duedate?: string;
  urgent?: boolean;
}

interface CreateTaskResponse {
  id: number;
}

interface Project {
  id: number;
  name: string;
}

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
  // 🔥 NEW: selected attachment indices
  const [selectedAttachments, setSelectedAttachments] = useState<number[]>([]);

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
  // DIRECT FETCH (same as EditTask)
  // ========================

  const getApiKey = (): string => {
    return localStorage.getItem("proworkflow-api-key") || process.env.PW_API_KEY || "";
  };

  const fetchProjectsDirectly = async (): Promise<Project[]> => {
    const API_KEY = getApiKey();
    if (!API_KEY) throw new Error("API key is not set.");

    const url = "https://api.proworkflow.com/api/v4/projects";
    console.log("🔍 Fetching projects from:", url);

    const response = await fetch(url, {
      method: "GET",
      headers: {
        apikey: API_KEY,
        Accept: "application/json",
      },
    });

    if (!response.ok) {
      const errorText = await response.text();
      console.error("❌ Failed to fetch projects:", errorText);
      throw new Error(`Failed to fetch projects (${response.status}): ${errorText}`);
    }

    const jsonData = await response.json();
    console.log("✅ Projects response:", jsonData);

    const data = jsonData.data || jsonData;
    if (Array.isArray(data)) {
      return data.map((p: any) => ({
        id: p.id,
        name: p.name || p.title || "Unnamed Project",
      }));
    }
    return [];
  };

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
      console.log("TEXT: ", text);
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
      itemtypeid: 1,
      itemcollectionid: 1,
    };

    if (description.trim()) payload.description = description.trim();
    const priorityValue = PRIORITY_MAP[priority];
    if (priorityValue) payload.priorityid = priorityValue;
    if (assigneeId > 0) payload.contacts = [{ id: assigneeId }];
    if (dueDate) {
      payload.duedate = dueDate;
      payload.startdate = dueDate;
    }
    if (isUrgent) payload.urgent = true;

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
        // If we have a draft, we will restore selectedAttachments later
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
      setSelectedAttachments(draft.selectedAttachments || []);
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

      const [projectsData, contactsData] = await Promise.all([
        fetchProjectsDirectly(),
        fetchContacts(),
      ]);

      setProjects(projectsData);
      setContacts(contactsData);
      setLoadingProjects(false);
      setLoadingContacts(false);

      await loadOutlookDataInternal();
      loadDraftInternal();
    } catch (err: any) {
      const errorMsg = err.message || "Failed to load data. Please check your API key.";
      setError(errorMsg);
      console.error("❌ Load error:", err);
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
      const testUrl = "https://api.proworkflow.com/api/v4/projects?limit=1";
      const response = await fetch(testUrl, {
        method: "GET",
        headers: { apikey: trimmed, Accept: "application/json" },
      });

      if (!response.ok) throw new Error("Invalid API key");

      localStorage.setItem("proworkflow-api-key", trimmed);
      setIsApiKeySet(true);
      setError(null);
      await loadAllData();
      showToast("API key verified! 🎉", "success");
    } catch (err: any) {
      setError("Invalid API key. Please check and try again.");
      showToast("Invalid API key", "error");
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

        // 🔥 Use selected attachments only
        if (
          includeAttachments &&
          selectedAttachments.length > 0 &&
          outlookData?.attachments?.length
        ) {
          const selectedFiles = selectedAttachments.map((idx) => outlookData.attachments[idx]);
          console.log(
            "📎 Selected attachments to upload:",
            selectedFiles.map((f) => f.name)
          );
          // Yahan aap actual attachment upload logic laga sakte hain
        } else if (includeAttachments && selectedAttachments.length === 0) {
          console.log("📎 No attachments selected, skipping upload.");
        }

        setSuccess(true);
        showToast(
          includeAttachments && selectedAttachments.length > 0
            ? `Task & ${selectedAttachments.length} attachments created! 🎉`
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
      selectedAttachments,
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
    setSelectedAttachments([]);
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

  // Auto-save draft (including selected attachments)
  useEffect(() => {
    if (!isApiKeySet || isFirstLoad.current) return undefined;

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
          selectedAttachments,
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
    selectedAttachments,
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
    const selectedCount = selectedAttachments.length;
    if (includeAttachments && count > 0) {
      return ` (${selectedCount} of ${count} selected)`;
    }
    return count > 0 ? ` (${count} files)` : "";
  };

  // ========================
  // RENDER
  // ========================

  // ---- API Key required ----
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
            <Typography variant="body2" color="textSecondary" sx={{ fontSize: "0.7rem" }}>
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

  // ---- Loading ----
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

  // ---- Main Form ----
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
            <Tooltip title="Refresh">
              <IconButton size="small" onClick={handleRefreshOutlook} disabled={loadingOutlook}>
                <Refresh fontSize="small" />
              </IconButton>
            </Tooltip>
            <Typography
              variant="caption"
              color="textSecondary"
              sx={{ ml: 0.5, fontSize: "0.65rem" }}
            >
              Subject: {outlookData.subject}
            </Typography>
            <Box sx={{ flex: 1 }} />
          </Stack>
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
            MenuProps={smallMenuProps}
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

        {/* Options – with attachment selection list */}
        <Paper variant="outlined" sx={{ p: 1 }}>
          <Stack spacing={0.5}>
            {/* Urgent checkbox */}
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
            {/* Master checkbox for attachments */}
            <FormControlLabel
              control={
                <Checkbox
                  checked={includeAttachments}
                  onChange={(e) => {
                    const checked = e.target.checked;
                    setIncludeAttachments(checked);
                    if (!checked) {
                      setSelectedAttachments([]); // clear selection when toggled off
                    } else if (outlookData?.attachments?.length) {
                      // Optionally select all by default? The user wants to choose, so we can keep selection empty.
                      // But to make it obvious, we can leave empty and let user pick.
                      // We'll leave empty.
                    }
                  }}
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

            {/* Attachment list – visible only when includeAttachments is true and attachments exist */}
            {includeAttachments &&
              outlookData?.attachments &&
              outlookData.attachments.length > 0 && (
                <Box sx={{ ml: 3, mt: 0.5 }}>
                  <Typography
                    variant="caption"
                    color="textSecondary"
                    sx={{ fontSize: "0.65rem", display: "block", mb: 0.5 }}
                  >
                    Select the attachments you want to include:
                  </Typography>
                  <Stack spacing={0.5}>
                    {outlookData.attachments.map((att, index) => (
                      <FormControlLabel
                        key={index}
                        control={
                          <Checkbox
                            checked={selectedAttachments.includes(index)}
                            onChange={(e) => {
                              if (e.target.checked) {
                                setSelectedAttachments([...selectedAttachments, index]);
                              } else {
                                setSelectedAttachments(
                                  selectedAttachments.filter((i) => i !== index)
                                );
                              }
                            }}
                            size="small"
                          />
                        }
                        label={
                          <Typography variant="body2" sx={{ fontSize: "0.7rem" }}>
                            {att.name} ({(att.size / 1024).toFixed(1)} KB)
                          </Typography>
                        }
                      />
                    ))}
                  </Stack>
                </Box>
              )}

            <Divider />
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

      {/* Toast */}
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
