import React, { useState, useEffect } from "react";
import {
  Box,
  Typography,
  FormControl,
  InputLabel,
  Select,
  MenuItem,
  TextField,
  Button,
  Alert,
  CircularProgress,
  Stack,
  Paper,
  Chip,
  Divider,
  Checkbox,
  FormControlLabel,
  InputAdornment,
  IconButton,
  Tooltip,
  Grid,
  List,
  ListItem,
  ListItemText,
  ListItemIcon,
  Avatar,
  Badge,
  ListItemButton,
  Card,
  CardContent,
  Skeleton,
  Dialog,
  DialogTitle,
  DialogContent,
  DialogContentText,
  DialogActions,
  Snackbar,
} from "@mui/material";
import {
  Edit,
  Search,
  Refresh,
  Save,
  Cancel,
  Folder,
  Person,
  PriorityHigh,
  Schedule,
  Description,
  Title,
  TaskAlt,
  Delete,
  History,
  CheckCircle,
  Warning,
  Info,
  Assignment,
  Event,
  PersonOutlined,
} from "@mui/icons-material";
import { proWorkflowApi, Project, User, Task, setApiKey } from "../../services/proworkflow";
import ApiTest from "../ApiTest";

// Helper function to get status color

// Helper function to get status color - with null/undefined handling
const getStatusColor = (status?: string): string => {
  switch (status?.toLowerCase()) {
    case "done":
    case "completed":
      return "#4caf50";
    case "in progress":
      return "#2196f3";
    case "in review":
      return "#ff9800";
    case "not started":
      return "#9e9e9e";
    case "blocked":
      return "#f44336";
    default:
      return "#9e9e9e";
  }
};

// Helper function to get status label - with null/undefined handling
const getStatusLabel = (status?: string): string => {
  switch (status?.toLowerCase()) {
    case "done":
    case "completed":
      return "Done";
    case "in progress":
      return "In Progress";
    case "in review":
      return "In Review";
    case "not started":
      return "Not Started";
    case "blocked":
      return "Blocked";
    default:
      return status || "Not Set";
  }
};

// Helper function to get priority color - with null/undefined handling
const getPriorityColor = (priority?: string): string => {
  switch (priority?.toLowerCase()) {
    case "high":
      return "#f44336";
    case "medium":
      return "#ff9800";
    case "low":
      return "#4caf50";
    default:
      return "#9e9e9e";
  }
};

// Helper function to get priority label - with null/undefined handling
const getPriorityLabel = (priority?: string): string => {
  switch (priority?.toLowerCase()) {
    case "high":
      return "High";
    case "medium":
      return "Medium";
    case "low":
      return "Low";
    default:
      return priority || "Not Set";
  }
};

const EditTask: React.FC = () => {
  // Selection state
  const [projects, setProjects] = useState<Project[]>([]);
  const [users, setUsers] = useState<User[]>([]);
  const [tasks, setTasks] = useState<Task[]>([]);
  const [selectedProjectId, setSelectedProjectId] = useState<number>(0);
  const [selectedTaskId, setSelectedTaskId] = useState<number>(0);

  // Task edit state
  const [taskName, setTaskName] = useState("");
  const [assigneeId, setAssigneeId] = useState<number>(0);
  const [description, setDescription] = useState("");
  const [priority, setPriority] = useState<"Low" | "Medium" | "High">("Medium");
  const [dueDate, setDueDate] = useState("");
  const [isUrgent, setIsUrgent] = useState(false);
  const [status, setStatus] = useState("");
  const [originalTask, setOriginalTask] = useState<Task | null>(null);

  // UI state
  const [loading, setLoading] = useState(true);
  const [loadingTasks, setLoadingTasks] = useState(false);
  const [saving, setSaving] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [success, setSuccess] = useState(false);
  const [isEditMode, setIsEditMode] = useState(false);
  const [searchTerm, setSearchTerm] = useState("");
  const [isApiKeySet, setIsApiKeySet] = useState(false);
  const [toastOpen, setToastOpen] = useState(false);
  const [toastMessage, setToastMessage] = useState("");
  const [toastSeverity, setToastSeverity] = useState<"success" | "error" | "info" | "warning">(
    "info"
  );
  const [showDeleteDialog, setShowDeleteDialog] = useState(false);
  const [deleting, setDeleting] = useState(false);

  // Load API key and initial data
  useEffect(() => {
    const initializeApp = async () => {
      const savedApiKey = localStorage.getItem("proworkflow-api-key");
      console.log("Edit Task Api Key", savedApiKey);

      if (savedApiKey && savedApiKey.trim()) {
        setApiKey(savedApiKey.trim());
        setIsApiKeySet(true);
        await loadData();
      } else {
        setIsApiKeySet(false);
        setError("Please set your ProWorkflow API key in the 'New Task' tab");
        setLoading(false);
      }
    };

    initializeApp();
  }, []);

  const loadData = async () => {
    setLoading(true);
    setError(null);
    try {
      const [projectsData, usersData] = await Promise.all([
        proWorkflowApi.getProjects(),
        proWorkflowApi.getUsers(),
      ]);
      console.log("user data", usersData);
      console.log("projedtData", projectsData);

      setProjects(projectsData);
      setUsers(usersData);
    } catch (err: any) {
      setError(err.message || "Failed to load data");
      console.error(err);
    } finally {
      setLoading(false);
    }
  };

  const loadTasks = async (projectId: number) => {
    if (!projectId) {
      setTasks([]);
      return;
    }

    setLoadingTasks(true);
    setError(null);
    try {
      const tasksData = await proWorkflowApi.getTasks(projectId);
      setTasks(tasksData);
      setSelectedTaskId(0);
      setIsEditMode(false);
      resetForm();

      if (tasksData.length === 0) {
        setError("No tasks found in this project");
      }
    } catch (err: any) {
      setError(err.message || "Failed to load tasks");
      console.error(err);
    } finally {
      setLoadingTasks(false);
    }
  };

  const handleProjectChange = (projectId: number) => {
    setSelectedProjectId(projectId);
    setSelectedTaskId(0);
    setTasks([]);
    setIsEditMode(false);
    resetForm();
    if (projectId) {
      loadTasks(projectId);
    }
  };

  const handleTaskSelect = (taskId: number) => {
    setSelectedTaskId(taskId);
    setIsEditMode(true);

    const task = tasks.find((t) => t.id === taskId);
    if (task) {
      setOriginalTask(task);
      setTaskName(task.name || "");
      setAssigneeId(task.assignee || 0);
      setDescription(task.description || "");
      setPriority((task.priority as "Low" | "Medium" | "High") || "Medium");
      setDueDate(task.duedate || "");
      setIsUrgent(task.urgent || false);
      setStatus(task.status || "");
      setError(null);
    } else {
      setError("Task not found");
      setIsEditMode(false);
    }
  };

  const resetForm = () => {
    setTaskName("");
    setAssigneeId(0);
    setDescription("");
    setPriority("Medium");
    setDueDate("");
    setIsUrgent(false);
    setStatus("");
    setOriginalTask(null);
    setError(null);
    setSuccess(false);
  };

  const handleSave = async () => {
    if (!selectedTaskId) {
      setError("No task selected");
      return;
    }

    if (!taskName.trim()) {
      setError("Task name is required");
      return;
    }

    setSaving(true);
    setError(null);

    try {
      const updatedData = {
        name: taskName.trim(),
        assignee: assigneeId || undefined,
        description: description.trim() || undefined,
        priority: priority,
        duedate: dueDate || undefined,
        urgent: isUrgent,
        status: status || undefined,
      };

      const result = await proWorkflowApi.updateTask(selectedTaskId, updatedData);
      console.log("Task updated:", result);

      const updatedTasks = tasks.map((task) =>
        task.id === selectedTaskId ? { ...task, ...updatedData } : task
      );
      setTasks(updatedTasks);
      setOriginalTask(result);

      setSuccess(true);
      showToast("Task updated successfully! ✅", "success");
      setTimeout(() => setSuccess(false), 5000);
    } catch (err: any) {
      setError(err.message || "Failed to update task");
      showToast(err.message || "Failed to update task", "error");
      console.error(err);
    } finally {
      setSaving(false);
    }
  };

  const handleDelete = async () => {
    if (!selectedTaskId) return;

    setDeleting(true);
    try {
      await proWorkflowApi.deleteTask(selectedTaskId);
      showToast("Task deleted successfully! 🗑️", "success");
      await loadTasks(selectedProjectId);
      setShowDeleteDialog(false);
      setIsEditMode(false);
      resetForm();
    } catch (err: any) {
      showToast(err.message || "Failed to delete task", "error");
    } finally {
      setDeleting(false);
    }
  };

  const handleCancel = () => {
    resetForm();
    setIsEditMode(false);
    setSelectedTaskId(0);
    if (selectedProjectId) {
      loadTasks(selectedProjectId);
    }
  };

  const handleRefresh = () => {
    if (selectedProjectId) {
      loadTasks(selectedProjectId);
    } else {
      loadData();
    }
    showToast("Data refreshed", "info");
  };

  const showToast = (message: string, severity: "success" | "error" | "info" | "warning") => {
    setToastMessage(message);
    setToastSeverity(severity);
    setToastOpen(true);
  };

  const handleToastClose = () => {
    setToastOpen(false);
  };

  const filteredTasks = tasks.filter(
    (task) =>
      task.name.toLowerCase().includes(searchTerm.toLowerCase()) ||
      (task.description && task.description.toLowerCase().includes(searchTerm.toLowerCase()))
  );

  const getUserName = (userId: number): string => {
    const user = users.find((u) => u.id === userId);
    return user ? user.name : "Unassigned";
  };

  // Show API Key Setup message
  if (!isApiKeySet && !loading) {
    return (
      <Box sx={{ p: 3 }}>
        <Paper sx={{ p: 4, textAlign: "center" }}>
          <Typography variant="h6" gutterBottom sx={{ fontWeight: 600 }}>
            🔑 ProWorkflow API Key Required
          </Typography>
          <Typography variant="body2" color="textSecondary" sx={{ mb: 3 }}>
            Please set your API key in the <strong>"New Task"</strong> tab first.
          </Typography>
          <Button
            variant="contained"
            onClick={() => {
              const tabs = document.querySelectorAll('[role="tab"]');
              if (tabs.length > 0) {
                (tabs[0] as HTMLElement).click();
              }
            }}
          >
            Go to New Task
          </Button>
        </Paper>
      </Box>
    );
  }

  if (loading) {
    return (
      <Box sx={{ p: 2 }}>
        <Stack spacing={3}>
          <Skeleton variant="rectangular" height={56} animation="wave" />
          <Skeleton variant="rectangular" height={200} animation="wave" />
        </Stack>
      </Box>
    );
  }

  return (
    <Box sx={{ maxWidth: "100%" }}>
      {/* Header */}
      <Stack direction="row" spacing={2} sx={{ alignItems: "center", mb: 3 }}>
        <Assignment color="primary" />
        <Typography variant="h6" sx={{ flex: 1 }}>
          Edit Existing Task
        </Typography>
        <Tooltip title="Refresh data">
          <IconButton onClick={handleRefresh} size="small">
            <Refresh />
          </IconButton>
        </Tooltip>
      </Stack>

      <Typography variant="body2" color="textSecondary" sx={{ mb: 3 }}>
        Select a project, then a task to view and edit its details.
      </Typography>

      {/* Error/Success Alerts */}
      {error && (
        <Alert severity="error" sx={{ mb: 2 }} onClose={() => setError(null)}>
          {error}
        </Alert>
      )}
      {success && (
        <Alert severity="success" sx={{ mb: 2 }} onClose={() => setSuccess(false)}>
          <Stack direction="row" sx={{ alignItems: "center" }} spacing={1}>
            <CheckCircle />
            <span>Task updated successfully! ✅</span>
          </Stack>
        </Alert>
      )}

      <Stack spacing={3}>
        {/* Project Selection */}
        <FormControl fullWidth>
          <InputLabel>Select Project</InputLabel>
          <Select
            value={selectedProjectId}
            onChange={(e) => handleProjectChange(e.target.value as number)}
            label="Select Project"
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

        {/* Task Selection - Simple List */}
        {selectedProjectId > 0 && (
          <Paper variant="outlined" sx={{ p: 2, bgcolor: "#fafafa" }}>
            <Stack spacing={2}>
              <Stack
                direction="row"
                spacing={2}
                sx={{ alignItems: "center", pb: 1, borderBottom: "1px solid #e0e0e0" }}
              >
                <Typography variant="subtitle2" sx={{ fontWeight: 600 }}>
                  Available Tasks
                </Typography>
                <Chip
                  label={tasks.length}
                  size="small"
                  color="primary"
                  variant="filled"
                  sx={{ fontWeight: 600 }}
                />
                <Box sx={{ flex: 1 }} />
                {loadingTasks && <CircularProgress size={20} />}
              </Stack>

              <TextField
                fullWidth
                size="small"
                placeholder="Search tasks..."
                value={searchTerm}
                onChange={(e) => setSearchTerm(e.target.value)}
                sx={{ "& .MuiOutlinedInput-root": { bgcolor: "white" } }}
                slotProps={{
                  input: {
                    startAdornment: (
                      <InputAdornment position="start">
                        <Search fontSize="small" color="action" />
                      </InputAdornment>
                    ),
                  },
                }}
              />

              {!loadingTasks && filteredTasks.length > 0 ? (
                <Box sx={{ maxHeight: "200px", overflow: "auto" }}>
                  {filteredTasks.map((task) => (
                    <Card
                      key={task.id}
                      variant="outlined"
                      sx={{
                        mb: 1,
                        cursor: "pointer",
                        borderColor: selectedTaskId === task.id ? "primary.main" : "#e0e0e0",
                        bgcolor: selectedTaskId === task.id ? "primary.light" : "white",
                        transition: "all 0.2s ease",
                        "&:hover": {
                          borderColor: "primary.main",
                          boxShadow: "0 2px 8px rgba(0,0,0,0.08)",
                        },
                      }}
                      onClick={() => handleTaskSelect(task.id)}
                    >
                      <CardContent sx={{ p: 2, "&:last-child": { pb: 2 } }}>
                        <Stack direction="row" spacing={2} sx={{ alignItems: "center" }}>
                          <Avatar
                            sx={{
                              width: 36,
                              height: 36,
                              bgcolor: getPriorityColor(task.priority || ""),
                              flexShrink: 0,
                            }}
                          >
                            {task.priority === "High" ? (
                              <PriorityHigh fontSize="small" sx={{ color: "white" }} />
                            ) : (
                              <TaskAlt fontSize="small" sx={{ color: "white" }} />
                            )}
                          </Avatar>

                          <Box sx={{ flex: 1, minWidth: 0 }}>
                            <Typography
                              variant="body2"
                              sx={{ fontWeight: selectedTaskId === task.id ? 600 : 500 }}
                            >
                              {task.name}
                            </Typography>
                            <Stack
                              direction="row"
                              spacing={1}
                              sx={{ mt: 0.5, alignItems: "center", flexWrap: "wrap" }}
                            >
                              <Chip
                                label={getStatusLabel(task.status)}
                                size="small"
                                sx={{
                                  bgcolor: getStatusColor(task.status || ""),
                                  color: "white",
                                  fontSize: "0.6rem",
                                  height: 20,
                                }}
                              />
                              {task.urgent && (
                                <Chip
                                  label="Urgent"
                                  size="small"
                                  color="error"
                                  sx={{ fontSize: "0.6rem", height: 20 }}
                                />
                              )}
                              {task.assignee && (
                                <Typography variant="caption" color="textSecondary">
                                  Assigned to: {getUserName(task.assignee)}
                                </Typography>
                              )}
                              {task.duedate && (
                                <Typography variant="caption" color="textSecondary">
                                  Due: {task.duedate}
                                </Typography>
                              )}
                            </Stack>
                          </Box>

                          {selectedTaskId === task.id && (
                            <CheckCircle color="primary" sx={{ fontSize: 20, flexShrink: 0 }} />
                          )}
                        </Stack>
                      </CardContent>
                    </Card>
                  ))}
                </Box>
              ) : (
                !loadingTasks && (
                  <Box sx={{ py: 3, textAlign: "center" }}>
                    <Info color="action" sx={{ fontSize: 36, opacity: 0.5, mb: 1 }} />
                    <Typography variant="body2" color="textSecondary">
                      {searchTerm
                        ? `No tasks match "${searchTerm}"`
                        : "No tasks available in this project"}
                    </Typography>
                  </Box>
                )
              )}
            </Stack>
          </Paper>
        )}

        {/* Edit Form - Clean and Simple */}
        {isEditMode && originalTask && (
          <>
            <Divider />
            <Paper elevation={1} sx={{ p: 3 }}>
              <Stack spacing={3}>
                {/* Task Info Header */}
                <Stack direction="row" spacing={2} sx={{ alignItems: "center", flexWrap: "wrap" }}>
                  <Badge color="primary" variant="dot">
                    <Typography variant="subtitle1" sx={{ fontWeight: 600 }}>
                      Editing: {originalTask.name}
                    </Typography>
                  </Badge>
                  <Box sx={{ flex: 1 }} />
                  <Chip
                    icon={<History />}
                    label={`ID: #${originalTask.id}`}
                    size="small"
                    variant="outlined"
                  />
                  <Button
                    variant="outlined"
                    color="error"
                    size="small"
                    startIcon={<Delete />}
                    onClick={() => setShowDeleteDialog(true)}
                  >
                    Delete
                  </Button>
                </Stack>

                <Divider />

                {/* Status */}
                <FormControl fullWidth>
                  <InputLabel>Status</InputLabel>
                  <Select value={status} onChange={(e) => setStatus(e.target.value)} label="Status">
                    <MenuItem value="">Not Set</MenuItem>
                    <MenuItem value="Not Started">Not Started</MenuItem>
                    <MenuItem value="In Progress">In Progress</MenuItem>
                    <MenuItem value="In Review">In Review</MenuItem>
                    <MenuItem value="Done">Done</MenuItem>
                    <MenuItem value="Blocked">Blocked</MenuItem>
                  </Select>
                </FormControl>

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
                  rows={3}
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

                {/* Priority & Due Date */}
                <Grid container spacing={2}>
                  <Grid size={{ xs: 12, sm: 6 }}>
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
                  </Grid>
                  <Grid size={{ xs: 12, sm: 6 }}>
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
                  </Grid>
                </Grid>

                {/* Urgent Checkbox */}
                <FormControlLabel
                  control={
                    <Checkbox
                      checked={isUrgent}
                      onChange={(e) => setIsUrgent(e.target.checked)}
                      color="error"
                    />
                  }
                  label={
                    <Stack direction="row" sx={{ alignItems: "center" }} spacing={1}>
                      <Warning color="error" fontSize="small" />
                      <Typography variant="body2">Mark as Urgent</Typography>
                    </Stack>
                  }
                />

                {/* Action Buttons */}
                <Stack direction="row" spacing={2} sx={{ justifyContent: "flex-end" }}>
                  <Button
                    variant="outlined"
                    onClick={handleCancel}
                    disabled={saving}
                    startIcon={<Cancel />}
                  >
                    Cancel
                  </Button>
                  <Button
                    variant="contained"
                    onClick={handleSave}
                    disabled={saving || !taskName.trim()}
                    startIcon={saving ? <CircularProgress size={20} /> : <Save />}
                  >
                    {saving ? "Saving..." : "Save Changes"}
                  </Button>
                </Stack>
              </Stack>
            </Paper>
          </>
        )}

        {/* Empty State */}
        {selectedProjectId === 0 && (
          <Alert severity="info" icon={<Info />}>
            Please select a project to view and edit tasks.
          </Alert>
        )}
      </Stack>

      {/* Delete Confirmation Dialog */}
      <Dialog open={showDeleteDialog} onClose={() => setShowDeleteDialog(false)}>
        <DialogTitle>Delete Task</DialogTitle>
        <DialogContent>
          <DialogContentText>
            Are you sure you want to delete <strong>"{originalTask?.name}"</strong>? This action
            cannot be undone.
          </DialogContentText>
        </DialogContent>
        <DialogActions>
          <Button onClick={() => setShowDeleteDialog(false)}>Cancel</Button>
          <Button
            onClick={handleDelete}
            color="error"
            variant="contained"
            disabled={deleting}
            startIcon={deleting ? <CircularProgress size={20} /> : <Delete />}
          >
            {deleting ? "Deleting..." : "Delete"}
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
      <ApiTest />
    </Box>
  );
};

export default EditTask;
