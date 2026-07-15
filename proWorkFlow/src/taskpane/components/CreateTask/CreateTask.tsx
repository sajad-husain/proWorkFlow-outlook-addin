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
  proWorkflowApi,
  Project,
  User,
  CreateTaskRequest,
  setApiKey,
} from "../../services/proworkflow";
import {
  getOutlookItemDataAsync,
  OutlookItemData,
  generateTaskDescription,
  cleanEmailBody,
} from "../../services/outlook";
import { useLocalStorage } from "../../hooks/useLocalStorage";

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

// 🔥 NEW: Helper to convert Base64 to Blob (if Outlook returns base64)
const base64ToBlob = (base64: string, mimeType: string = "application/octet-stream"): Blob => {
  const byteCharacters = atob(base64);
  const byteNumbers = new Array(byteCharacters.length);
  for (let i = 0; i < byteCharacters.length; i++) {
    byteNumbers[i] = byteCharacters.charCodeAt(i);
  }
  const byteArray = new Uint8Array(byteNumbers);
  return new Blob([byteArray], { type: mimeType });
};

const CreateTask: React.FC = () => {
  // Form state
  const [taskName, setTaskName] = useState("");
  const [projectId, setProjectId] = useState<number>(0);
  const [assigneeId, setAssigneeId] = useState<number>(0);
  const [description, setDescription] = useState("");
  const [priority, setPriority] = useState<"Low" | "Medium" | "High">("Medium");
  const [dueDate, setDueDate] = useState("");
  const [isUrgent, setIsUrgent] = useState(false);
  const [includeAttachments, setIncludeAttachments] = useState(false);

  // Data state
  const [projects, setProjects] = useState<Project[]>([]);
  const [users, setUsers] = useState<User[]>([]);
  const [loadingProjects, setLoadingProjects] = useState(false);
  const [loadingUsers, setLoadingUsers] = useState(false);

  // UI state
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
  const isInitialized = useRef(false);

  // Auto-save draft
  const [draft, setDraft] = useLocalStorage<DraftData | null>("proworkflow-task-draft", null);

  // Load Outlook data - defined first so it can be used in loadAllData
  const loadOutlookDataInternal = useCallback(async () => {
    setLoadingOutlook(true);
    try {
      const data = await getOutlookItemDataAsync();
      setOutlookData(data);
      if (data) {
        if (!draft?.taskName) {
          setTaskName(data.subject || "");
        }
        if (!draft?.description) {
          const generatedDescription = generateTaskDescription(data);
          setDescription(generatedDescription);
        }
      }
    } catch (err) {
      console.warn("Could not load Outlook data:", err);
    } finally {
      setLoadingOutlook(false);
    }
  }, [draft]);

  // Load draft
  const loadDraftInternal = useCallback(() => {
    if (draft) {
      setTaskName(draft.taskName || "");
      setProjectId(draft.projectId || 0);
      setAssigneeId(draft.assigneeId || 0);
      setDescription(draft.description || "");
      setPriority(draft.priority || "Medium");
      setDueDate(draft.dueDate || "");
      setIsUrgent(draft.isUrgent || false);
      setIncludeAttachments(draft.includeAttachments || false);
    }
  }, [draft]);

  // Load all data
  const loadAllData = useCallback(async () => {
    setLoadingData(true);
    setError(null);
    try {
      setLoadingProjects(true);
      setLoadingUsers(true);

      const [projectsData, usersData] = await Promise.all([
        proWorkflowApi.getProjects(),
        proWorkflowApi.getUsers(),
      ]);

      setProjects(projectsData);
      setUsers(usersData);
      setLoadingProjects(false);
      setLoadingUsers(false);

      await loadOutlookDataInternal();
      loadDraftInternal();
    } catch (err: any) {
      const errorMsg = err.message || "Failed to load data. Please check your API key.";
      setError(errorMsg);
      console.error(err);
      setLoadingProjects(false);
      setLoadingUsers(false);
    } finally {
      setLoadingData(false);
    }
  }, [loadOutlookDataInternal, loadDraftInternal]);

  // Initialize app - runs only once
  useEffect(() => {
    if (isInitialized.current) return;
    isInitialized.current = true;

    const initializeApp = async () => {
      // 🔥 .env fallback (already added)
      const defaultKey = process.env.PW_API_KEY || "";
      const savedKey = localStorage.getItem("proworkflow-api-key");
      const finalKey = savedKey && savedKey.trim() ? savedKey.trim() : defaultKey;

      if (finalKey) {
        setApiKey(finalKey);
        setIsApiKeySet(true);
        setApiKeyInput(finalKey);
        setError(null);
        await loadAllData();
      } else {
        setIsApiKeySet(false);
        setLoadingData(false);
      }
    };

    initializeApp();
  }, [loadAllData]);

  // Auto-save draft on form change - with debounce
  useEffect((): void | (() => void) => {
    if (!isApiKeySet) return;

    const timeoutId = setTimeout(() => {
      if (taskName || description || projectId || assigneeId) {
        const draftData: DraftData = {
          taskName,
          projectId,
          assigneeId,
          description,
          priority,
          dueDate,
          isUrgent,
          includeAttachments,
        };
        setDraft(draftData);
      }
    }, 500);

    return () => clearTimeout(timeoutId);
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

  const handleRefreshOutlook = () => {
    loadOutlookDataInternal();
    showToast("Email data refreshed", "info");
  };

  // API Key setup
  const handleApiKeySubmit = async () => {
    if (!apiKeyInput || !apiKeyInput.trim()) {
      showToast("Please enter a valid API key", "error");
      return;
    }

    try {
      setLoading(true);
      setApiKey(apiKeyInput.trim());
      const isValid = await proWorkflowApi.testApiKey(apiKeyInput.trim());

      if (isValid) {
        localStorage.setItem("proworkflow-api-key", apiKeyInput.trim());
        setIsApiKeySet(true);
        setError(null);
        await loadAllData();
        showToast("API key verified successfully! 🎉", "success");
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
  };

  // Keyboard shortcuts
  useEffect(() => {
    const handleKeyDown = (e: KeyboardEvent) => {
      if ((e.ctrlKey || e.metaKey) && e.key === "Enter") {
        e.preventDefault();
        if (taskName.trim() && projectId) {
          handleSubmitClick();
        }
      }
      if (e.key === "Escape" && !showConfirmDialog) {
        e.preventDefault();
        setShowConfirmDialog(true);
        setPendingAction("reset");
      }
    };

    window.addEventListener("keydown", handleKeyDown);
    return () => window.removeEventListener("keydown", handleKeyDown);
  }, [taskName, projectId, showConfirmDialog]);

  const handleSubmitClick = () => {
    if (taskName.trim() && projectId) {
      setShowConfirmDialog(true);
      setPendingAction("submit");
    } else {
      showToast("Please fill in all required fields", "warning");
    }
  };

  const handleConfirmAction = () => {
    setShowConfirmDialog(false);
    if (pendingAction === "submit") {
      handleSubmit();
    } else if (pendingAction === "reset") {
      handleReset();
    }
    setPendingAction(null);
  };

  // 🔥 NEW: Main submit with attachment upload
  const handleSubmit = async (e?: React.FormEvent | React.SyntheticEvent) => {
    if (e && e.preventDefault) {
      e.preventDefault();
    }

    setError(null);
    setSuccess(false);

    // Validation
    if (!taskName.trim()) {
      showToast("Task name is required", "error");
      return;
    }
    if (!projectId) {
      showToast("Please select a project", "error");
      return;
    }

    setLoading(true);

    try {
      // 1. Create the task
      const taskData: CreateTaskRequest = {
        name: taskName.trim(),
        projectid: projectId,
        description: description.trim() || undefined,
        assignee: assigneeId || undefined,
        priority: priority,
        duedate: dueDate || undefined,
        urgent: isUrgent,
      };

      const result = await proWorkflowApi.createTask(taskData);
      console.log("Task created:", result);

      // Extract task ID (API might return { data: { id } } or just { id })
      const newTaskId = result.id || result?.data?.id;
      if (!newTaskId) {
        throw new Error("Task created but no ID returned");
      }

      // 2. 🔥 Upload attachments if checkbox is ticked and attachments exist
      if (includeAttachments && outlookData?.attachments && outlookData.attachments.length > 0) {
        console.log(`📎 Uploading ${outlookData.attachments.length} attachments...`);

        // Map each attachment to an upload promise
        const uploadPromises = outlookData.attachments.map(async (att) => {
          // Determine if att is already a File/Blob or base64 string
          let fileToUpload: File | Blob;
          let fileName = att.name || "attachment.bin";

          if (att instanceof File || att instanceof Blob) {
            // Already a File/Blob
            fileToUpload = att;
          } else if (typeof att === "string" && (att as string).startsWith("data:")) {
            // Data URL – convert to Blob
            const response = await fetch(att);
            fileToUpload = await response.blob();
            // Extract filename from data URL if possible, else use default
          } else if (typeof att === "string" && (att as string).length > 0) {
            // Assume it's base64 string (without data: prefix)
            // Need to know mime type, we can guess or let user pass in att.mimeType
            // For simplicity, assume application/octet-stream
            const mimeType = (att as any).mimeType || "application/octet-stream";
            fileToUpload = base64ToBlob(att, mimeType);
          } else {
            // Unknown format – skip or throw
            console.warn("Skipping attachment with unknown format:", att);
            return null;
          }

          // Actually upload
          // proWorkflowApi may not have uploadAttachment typed; cast to any to call if available
          return (proWorkflowApi as any).uploadAttachment
            ? (proWorkflowApi as any).uploadAttachment(newTaskId, fileToUpload, fileName)
            : Promise.reject(new Error("uploadAttachment not implemented on proWorkflowApi"));
        });

        // Wait for all uploads to finish
        const uploadResults = await Promise.all(
          uploadPromises.map((p) =>
            p.catch((err) => ({ status: "rejected" as const, reason: err }))
          )
        );
        const failed = uploadResults.filter((r) => (r as any).status === "rejected");
        if (failed.length > 0) {
          console.warn(`${failed.length} attachment(s) failed to upload.`);
          showToast(`Task created, but ${failed.length} attachment(s) failed.`, "warning");
        } else {
          console.log("✅ All attachments uploaded successfully.");
        }
      }

      setSuccess(true);
      showToast(
        includeAttachments && outlookData?.attachments?.length
          ? "Task and attachments created successfully! 🎉"
          : "Task created successfully! 🎉",
        "success"
      );
      setDraft(null);

      // Reset form after success (optional)
      setTimeout(() => {
        setSuccess(false);
        // Uncomment if you want auto-reset:
        // handleReset();
      }, 5000);
    } catch (err: any) {
      setError(err.message || "Failed to create task");
      showToast(err.message || "Failed to create task", "error");
      console.error(err);
    } finally {
      setLoading(false);
    }
  };

  const handleReset = () => {
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
      const generatedDescription = generateTaskDescription(outlookData);
      setDescription(generatedDescription);
    }
    showToast("Form reset successfully", "info");
  };

  const showToast = (message: string, severity: "success" | "error" | "info" | "warning") => {
    setToastMessage(message);
    setToastSeverity(severity);
    setToastOpen(true);
  };

  const handleToastClose = () => {
    setToastOpen(false);
  };

  const getAttachmentsLabel = (): string => {
    const count = outlookData?.attachments?.length ?? 0;
    return count > 0 ? ` (${count} files)` : " (No attachments found)";
  };

  // Show API Key Setup
  if (!isApiKeySet && !loadingData) {
    return (
      <Box sx={{ p: 3 }}>
        <Paper sx={{ p: 4, textAlign: "center" }}>
          <Typography variant="h5" gutterBottom sx={{ fontWeight: 600 }}>
            🔑 ProWorkflow API Key
          </Typography>
          <Typography variant="body2" color="textSecondary" sx={{ mb: 3 }}>
            Please enter your ProWorkflow API key to connect the add-in. You can find it in your
            ProWorkflow account settings.
          </Typography>

          <Box sx={{ maxWidth: 400, mx: "auto" }}>
            <TextField
              fullWidth
              label="API Key"
              placeholder="e.g., EIOM-ZTAY-VAD8-D2Y4-PWFO9JJ-AS11627"
              value={apiKeyInput}
              onChange={(e) => setApiKeyInput(e.target.value)}
              onKeyDown={(e) => {
                if (e.key === "Enter") {
                  handleApiKeySubmit();
                }
              }}
              sx={{ mb: 2 }}
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
              onClick={handleApiKeySubmit}
              disabled={loading || !apiKeyInput.trim()}
              startIcon={loading ? <CircularProgress size={20} /> : <CheckCircle />}
              size="large"
            >
              {loading ? "Verifying..." : "Verify & Connect"}
            </Button>

            {error && (
              <Alert severity="error" sx={{ mt: 2 }}>
                {error}
              </Alert>
            )}
          </Box>

          <Paper variant="outlined" sx={{ p: 2, mt: 3, bgcolor: "#f8f9fa", textAlign: "left" }}>
            <Typography variant="caption" color="textSecondary" sx={{ fontWeight: 600 }}>
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
        </Paper>
      </Box>
    );
  }

  if (loadingData) {
    return (
      <Box sx={{ p: 2 }}>
        <Stack spacing={3}>
          <Skeleton variant="rectangular" height={60} animation="wave" />
          <Skeleton variant="rectangular" height={56} animation="wave" />
          <Skeleton variant="rectangular" height={56} animation="wave" />
          <Skeleton variant="rectangular" height={56} animation="wave" />
          <Skeleton variant="rectangular" height={100} animation="wave" />
          <Skeleton variant="rectangular" height={56} animation="wave" />
          <Skeleton variant="rectangular" height={80} animation="wave" />
          <Skeleton variant="rectangular" height={40} animation="wave" />
        </Stack>
      </Box>
    );
  }

  return (
    <Box
      component="form"
      onSubmit={(e: React.FormEvent) => handleSubmit(e)}
      sx={{ maxWidth: "100%" }}
    >
      {/* Header with Outlook info */}
      {outlookData && (
        <Paper
          elevation={0}
          sx={{
            p: 2,
            mb: 3,
            bgcolor: "#f8f9fa",
            border: "1px solid #e0e0e0",
            borderRadius: 1,
          }}
        >
          <Stack direction="row" spacing={1} sx={{ alignItems: "center", flexWrap: "wrap", mb: 1 }}>
            <Email fontSize="small" color="primary" />
            <Typography variant="caption" color="textSecondary">
              From:
            </Typography>
            <Chip label={outlookData.sender} size="small" variant="outlined" color="primary" />
            <Typography variant="caption" color="textSecondary" sx={{ ml: 1 }}>
              Subject: {outlookData.subject}
            </Typography>
            <Box sx={{ flex: 1 }} />
            <Tooltip title="Refresh email data">
              <IconButton size="small" onClick={handleRefreshOutlook} disabled={loadingOutlook}>
                <Refresh fontSize="small" />
              </IconButton>
            </Tooltip>
            <Tooltip title="Toggle email preview">
              <IconButton size="small" onClick={() => setShowEmailPreview(!showEmailPreview)}>
                {showEmailPreview ? <ExpandLess /> : <ExpandMore />}
              </IconButton>
            </Tooltip>
          </Stack>

          <Collapse in={showEmailPreview}>
            <Divider sx={{ my: 1 }} />
            <Box sx={{ mt: 1 }}>
              <Typography variant="caption" color="textSecondary" sx={{ mb: 1, display: "block" }}>
                Email Preview:
              </Typography>
              <Paper
                variant="outlined"
                sx={{
                  p: 1,
                  bgcolor: "white",
                  maxHeight: "150px",
                  overflow: "auto",
                  fontSize: "0.875rem",
                  whiteSpace: "pre-wrap",
                  wordBreak: "break-word",
                }}
              >
                {cleanEmailBody(outlookData.body || "No content") || "No content"}
              </Paper>
              {outlookData.attachments && outlookData.attachments.length > 0 && (
                <Box sx={{ mt: 1 }}>
                  <Typography variant="caption" color="textSecondary">
                    Attachments: {outlookData.attachments.map((a) => a.name).join(", ")}
                  </Typography>
                </Box>
              )}
            </Box>
          </Collapse>
        </Paper>
      )}

      {!outlookData && !loadingOutlook && (
        <Alert severity="info" sx={{ mb: 2 }}>
          No email selected. Please select an email in Outlook to auto-fill task details.
          <Button size="small" onClick={handleRefreshOutlook} sx={{ ml: 2 }}>
            Refresh
          </Button>
        </Alert>
      )}

      {loadingOutlook && (
        <Box sx={{ display: "flex", alignItems: "center", mb: 2 }}>
          <CircularProgress size={20} sx={{ mr: 1 }} />
          <Typography variant="caption" color="textSecondary">
            Loading email data...
          </Typography>
        </Box>
      )}

      {/* Error/Success Alerts */}
      {error && (
        <Alert severity="error" sx={{ mb: 2 }} onClose={() => setError(null)}>
          {error}
        </Alert>
      )}
      {success && (
        <Alert severity="success" sx={{ mb: 2 }} icon={<CheckCircle />}>
          Task created successfully! 🎉
        </Alert>
      )}

      {/* Draft indicator */}
      {draft && (
        <Alert severity="info" sx={{ mb: 2 }} icon={<Warning />}>
          You have a saved draft. Continue editing or click Reset to clear.
        </Alert>
      )}

      <Stack spacing={3}>
        {/* Task Name */}
        <TextField
          required
          fullWidth
          label="Task Name"
          value={taskName}
          onChange={(e) => setTaskName(e.target.value)}
          placeholder="Enter task name..."
          slotProps={{
            input: {
              startAdornment: (
                <InputAdornment position="start">
                  <Title color="action" />
                </InputAdornment>
              ),
            },
          }}
        />

        {/* Project */}
        <FormControl fullWidth required>
          <InputLabel>Project</InputLabel>
          <Select
            value={projectId}
            onChange={(e) => setProjectId(e.target.value as number)}
            label="Project"
            disabled={loadingProjects}
            startAdornment={
              loadingProjects ? (
                <CircularProgress size={20} sx={{ ml: 1 }} />
              ) : (
                <Folder color="action" sx={{ ml: 1 }} />
              )
            }
          >
            <MenuItem value={0}>Select a project...</MenuItem>
            {projects.map((project) => (
              <MenuItem key={project.id} value={project.id}>
                {project.name}
              </MenuItem>
            ))}
          </Select>
        </FormControl>

        {/* Assignee */}
        <FormControl fullWidth>
          <InputLabel>Assignee</InputLabel>
          <Select
            value={assigneeId}
            onChange={(e) => setAssigneeId(e.target.value as number)}
            label="Assignee"
            disabled={loadingUsers}
            startAdornment={
              loadingUsers ? (
                <CircularProgress size={20} sx={{ ml: 1 }} />
              ) : (
                <Person color="action" sx={{ ml: 1 }} />
              )
            }
          >
            <MenuItem value={0}>Unassigned</MenuItem>
            {users.map((user) => (
              <MenuItem key={user.id} value={user.id}>
                {user.name}
              </MenuItem>
            ))}
          </Select>
        </FormControl>

        {/* Description */}
        <TextField
          fullWidth
          multiline
          rows={4}
          label="Description"
          value={description}
          onChange={(e) => setDescription(e.target.value)}
          placeholder="Task description..."
          slotProps={{
            input: {
              startAdornment: (
                <InputAdornment position="start">
                  <Description color="action" sx={{ mr: 1, mt: 1 }} />
                </InputAdornment>
              ),
            },
          }}
        />

        {/* Priority & Due Date Row */}
        <Stack direction="row" spacing={2}>
          <FormControl fullWidth>
            <InputLabel>Priority</InputLabel>
            <Select
              value={priority}
              onChange={(e) => setPriority(e.target.value as "Low" | "Medium" | "High")}
              label="Priority"
              startAdornment={<PriorityHigh color="action" sx={{ mr: 1 }} />}
            >
              <MenuItem value="Low">Low</MenuItem>
              <MenuItem value="Medium">Medium</MenuItem>
              <MenuItem value="High">High</MenuItem>
            </Select>
          </FormControl>

          <TextField
            fullWidth
            type="date"
            label="Due Date"
            value={dueDate}
            onChange={(e) => setDueDate(e.target.value)}
            slotProps={{
              input: {
                startAdornment: (
                  <InputAdornment position="start">
                    <Schedule color="action" sx={{ mr: 1 }} />
                  </InputAdornment>
                ),
              },
            }}
          />
        </Stack>

        {/* Options */}
        <Paper variant="outlined" sx={{ p: 2 }}>
          <Stack spacing={2}>
            <FormControlLabel
              control={
                <Checkbox
                  checked={isUrgent}
                  onChange={(e) => setIsUrgent(e.target.checked)}
                  color="error"
                />
              }
              label={
                <Typography variant="body2">
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
                  disabled={!outlookData?.attachments?.length}
                />
              }
              label={
                <Stack direction="row" spacing={1} sx={{ alignItems: "center" }}>
                  <AttachFile fontSize="small" />
                  <Typography variant="body2">
                    Include email attachments{getAttachmentsLabel()}
                  </Typography>
                </Stack>
              }
            />
          </Stack>
        </Paper>

        {/* Keyboard shortcuts hint */}
        <Typography variant="caption" color="textSecondary" sx={{ textAlign: "center" }}>
          <Stack
            direction="row"
            spacing={2}
            sx={{ justifyContent: "center", alignItems: "center", flexWrap: "wrap" }}
          >
            <span>⌘+Enter / Ctrl+Enter to submit</span>
            <span>•</span>
            <span>Esc to reset</span>
          </Stack>
        </Typography>

        {/* Action Buttons */}
        <Stack direction="row" spacing={2} sx={{ justifyContent: "flex-end" }}>
          <Button
            variant="outlined"
            onClick={() => {
              setShowConfirmDialog(true);
              setPendingAction("reset");
            }}
            disabled={loading}
            startIcon={<Cancel />}
          >
            Reset
          </Button>
          <Button
            type="button"
            variant="contained"
            disabled={loading || !taskName.trim() || !projectId}
            startIcon={loading ? <CircularProgress size={20} /> : <Send />}
            onClick={handleSubmitClick}
          >
            {loading ? "Creating..." : "Create Task"}
          </Button>
        </Stack>
      </Stack>

      {/* Confirmation Dialog */}
      <Dialog open={showConfirmDialog} onClose={() => setShowConfirmDialog(false)}>
        <DialogTitle>Confirm Action</DialogTitle>
        <DialogContent>
          <DialogContentText>
            {pendingAction === "submit"
              ? "Are you sure you want to create this task?"
              : "Are you sure you want to reset all form fields? This will clear your draft."}
          </DialogContentText>
        </DialogContent>
        <DialogActions>
          <Button onClick={() => setShowConfirmDialog(false)}>Cancel</Button>
          <Button
            onClick={handleConfirmAction}
            variant="contained"
            color={pendingAction === "submit" ? "primary" : "error"}
          >
            {pendingAction === "submit" ? "Create" : "Reset"}
          </Button>
        </DialogActions>
      </Dialog>

      {/* Toast Notification */}
      <Snackbar
        open={toastOpen}
        autoHideDuration={4000}
        onClose={handleToastClose}
        anchorOrigin={{ vertical: "bottom", horizontal: "center" }}
      >
        <Alert onClose={handleToastClose} severity={toastSeverity} variant="filled">
          {toastMessage}
        </Alert>
      </Snackbar>
    </Box>
  );
};

export default CreateTask;
