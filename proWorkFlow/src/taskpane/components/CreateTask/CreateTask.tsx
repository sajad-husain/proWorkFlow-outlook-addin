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
  IconButton,
  Tooltip,
  InputAdornment,
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
} from "@mui/icons-material";
import {
  proWorkflowApi,
  Project,
  User,
  CreateTaskRequest,
  setApiKey,
} from "../../services/proworkflow";
import { getOutlookItemDataAsync, OutlookItemData } from "../../services/outlook";

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
  const [error, setError] = useState<string | null>(null);
  const [success, setSuccess] = useState(false);
  const [outlookData, setOutlookData] = useState<OutlookItemData | null>(null);

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
    try {
      const data = await getOutlookItemDataAsync();
      setOutlookData(data);
      if (data) {
        setTaskName(data.subject || "");
        setDescription(data.body || "");
      }
    } catch (err) {
      console.warn("Could not load Outlook data:", err);
    }
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
      // Reset form (optional)
      // Reset form after success
      setTimeout(() => {
        setSuccess(false);
        // Don't reset if user wants to keep data
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
      setDescription(outlookData.body || "");
    }
  };

  if (loadingData) {
    return (
      <Box
        sx={{ display: "flex", justifyContent: "center", alignItems: "center", minHeight: "300px" }}
      >
        <CircularProgress />
        <Typography variant="body2" sx={{ ml: 2 }}>
          Loading data...
        </Typography>
      </Box>
    );
  }

  return (
    <Box component="form" onSubmit={handleSubmit} sx={{ maxWidth: "100%" }}>
      {/* Header with Outlook info */}
      {outlookData && (
        <Paper elevation={0} sx={{ p: 2, mb: 3, bgcolor: "#f8f9fa" }}>
          <Stack direction="row" spacing={1} sx={{ alignItems: "center", flexWrap: "wrap" }}>
            <Typography variant="caption" color="textSecondary">
              From:
            </Typography>
            <Chip label={outlookData.sender} size="small" variant="outlined" />
            <Typography variant="caption" color="textSecondary" sx={{ ml: 1 }}>
              Subject: {outlookData.subject}
            </Typography>
          </Stack>
        </Paper>
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
                    Include email attachments
                    {(() => {
                      const count = outlookData?.attachments?.length ?? 0;
                      if (count > 0) {
                        return ` (${count} files)`;
                      }
                      return " (No attachments found)";
                    })()}
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
    </Box>
  );
};

export default CreateTask;
