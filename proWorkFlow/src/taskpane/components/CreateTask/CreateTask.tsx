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
  Skeleton,
  Dialog,
  DialogTitle,
  DialogContent,
  DialogContentText,
  DialogActions,
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
  Download,
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
  // Loading Dialog
  const [showConfirmDialog, setShowConfirmDialog] = useState(false);
  const [pendingAction, setPendingAction] = useState<"submit" | "reset" | null>(null);

  // Add handlers
  const handleSubmitClick = () => {
    if (taskName.trim() && projectId) {
      setShowConfirmDialog(true);
      setPendingAction("submit");
    } else {
      handleSubmit(new Event("submit") as any);
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
  }, []);

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
        // Auto-fill task name from subject
        setTaskName(data.subject || "");

        // Auto-fill description from email
        const generatedDescription = generateTaskDescription(data);
        setDescription(generatedDescription);
      }
    } catch (err) {
      console.warn("Could not load Outlook data:", err);
      setError("Could not load email data. Please ensure you have an email selected.");
    } finally {
      setLoadingOutlook(false);
    }
  };

  const handleRefreshOutlook = () => {
    loadOutlookData();
  };

  const handleSubmit = async (e: React.FormEvent) => {
    e.preventDefault();
    setError(null);
    setSuccess(false);

    // Validation
    if (!taskName.trim()) {
      setError("Task name is required");
      return;
    }
    if (!projectId) {
      setError("Please select a project");
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
      setTimeout(() => {
        setSuccess(false);
      }, 5000);
    } catch (err: any) {
      setError(err.message || "Failed to create task");
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

    // Reload Outlook data if available
    if (outlookData) {
      setTaskName(outlookData.subject || "");
      const generatedDescription = generateTaskDescription(outlookData);
      setDescription(generatedDescription);
    }
  };

  // Helper function for attachments label
  const getAttachmentsLabel = (): string => {
    const count = outlookData?.attachments?.length ?? 0;
    return count > 0 ? ` (${count} files)` : " (No attachments found)";
  };

  if (loadingData) {
    return (
      <Box sx={{ p: 2 }}>
        <Stack spacing={3}>
          <Skeleton variant="rectangular" height={56} animation="wave" />
          <Skeleton variant="rectangular" height={56} animation="wave" />
          <Skeleton variant="rectangular" height={56} animation="wave" />
          <Skeleton variant="rectangular" height={100} animation="wave" />
          <Skeleton variant="rectangular" height={56} animation="wave" />
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
            <IconButton
              size="small"
              onClick={handleRefreshOutlook}
              disabled={loadingOutlook}
              title="Refresh email data"
            >
              <Refresh fontSize="small" />
            </IconButton>
            <IconButton
              size="small"
              onClick={() => setShowEmailPreview(!showEmailPreview)}
              title="Toggle email preview"
            >
              {showEmailPreview ? <ExpandLess /> : <ExpandMore />}
            </IconButton>
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
        <Alert severity="success" sx={{ mb: 2 }}>
          Task created successfully! ✅
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

        {/* Action Buttons */}
        <Stack direction="row" spacing={2} sx={{ justifyContent: "flex-end" }}>
          <Button
            variant="outlined"
            onClick={handleReset}
            disabled={loading}
            startIcon={<Cancel />}
          >
            Reset
          </Button>
          <Button
            type="submit"
            variant="contained"
            disabled={loading || !taskName.trim() || !projectId}
            startIcon={loading ? <CircularProgress size={20} /> : <Send />}
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
              : "Are you sure you want to reset all form fields?"}
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
    </Box>
  );
};

export default CreateTask;
