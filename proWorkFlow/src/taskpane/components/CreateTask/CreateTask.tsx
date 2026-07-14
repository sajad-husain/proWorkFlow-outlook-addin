import React, { useState, useEffect } from "react";
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

  // Auto-save draft
  const [draft, setDraft] = useLocalStorage<DraftData | null>("proworkflow-task-draft", null);
  const [apiKey, setApiKeyState] = useState<string>("");

  // Load initial data
  useEffect(() => {
    const testApiKey = "your-test-api-key-here";
    if (testApiKey && testApiKey !== "your-test-api-key-here") {
      setApiKey(testApiKey);
    } else {
      setError("Please configure your ProWorkflow API key");
    }
    loadInitialData();
    loadOutlookData();
    loadDraft();
  }, []);

  // Auto-save draft on form change
  useEffect(() => {
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
  }, [
    taskName,
    projectId,
    assigneeId,
    description,
    priority,
    dueDate,
    isUrgent,
    includeAttachments,
  ]);

  const loadDraft = () => {
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
  };

  const loadInitialData = async () => {
    setLoadingData(true);
    setError(null);
    try {
      const [projectsData, usersData] = await Promise.all([
        proWorkflowApi.getProjects(),
        proWorkflowApi.getUsers(),
      ]);
      setProjects(projectsData);
      setUsers(usersData);
    } catch (err) {
      setError("Failed to load data. Please check your API key.");
      console.error(err);
    } finally {
      setLoadingData(false);
    }
  };

  const loadOutlookData = async () => {
    setLoadingOutlook(true);
    try {
      const data = await getOutlookItemDataAsync();
      setOutlookData(data);
      if (data) {
        if (!draft || !draft.taskName) {
          setTaskName(data.subject || "");
        }
        if (!draft || !draft.description) {
          const generatedDescription = generateTaskDescription(data);
          setDescription(generatedDescription);
        }
      }
    } catch (err) {
      console.warn("Could not load Outlook data:", err);
    } finally {
      setLoadingOutlook(false);
    }
  };

  const handleRefreshOutlook = () => {
    loadOutlookData();
    showToast("Email data refreshed", "info");
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
      handleSubmit(new Event("submit") as any);
    } else if (pendingAction === "reset") {
      handleReset();
    }
    setPendingAction(null);
  };

  const handleSubmit = async (e: React.FormEvent) => {
    e.preventDefault();
    setError(null);
    setSuccess(false);

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

      setSuccess(true);
      showToast("Task created successfully! 🎉", "success");

      // Clear draft on success
      setDraft(null);

      setTimeout(() => {
        setSuccess(false);
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
    <Box component="form" onSubmit={handleSubmit} sx={{ maxWidth: "100%" }}>
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

          {/* Email Preview Collapse */}
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
            startAdornment={<Folder color="action" sx={{ mr: 1 }} />}
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
            startAdornment={<Person color="action" sx={{ mr: 1 }} />}
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
