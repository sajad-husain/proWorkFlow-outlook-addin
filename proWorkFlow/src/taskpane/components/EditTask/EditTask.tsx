import React, { useState, useEffect, useRef } from "react";
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
  Avatar,
  Badge,
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
} from "@mui/icons-material";
import { proWorkflowApi, Project, User, Task, setApiKey } from "../../services/proworkflow";
import { smallLabelStyles, smallMenuProps } from "../../theme/styles";

// Helper functions
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
  const [apiKeyInput, setApiKeyInput] = useState("");
  const [toastOpen, setToastOpen] = useState(false);
  const [toastMessage, setToastMessage] = useState("");
  const [toastSeverity, setToastSeverity] = useState<"success" | "error" | "info" | "warning">(
    "info"
  );
  const [showDeleteDialog, setShowDeleteDialog] = useState(false);
  const [deleting, setDeleting] = useState(false);
  const isInitialized = useRef(false);

  // ----- Direct fetch functions (bypassing service) -----

  // 🔥 Direct fetch for projects
  const fetchProjectsDirectly = async (): Promise<Project[]> => {
    const API_KEY = localStorage.getItem("proworkflow-api-key") || process.env.PW_API_KEY || "";
    if (!API_KEY) {
      throw new Error("API key is not set.");
    }

    const url = "https://api.proworkflow.com/api/v4/projects";
    console.log("🔍 Direct fetch projects from:", url);

    const response = await fetch(url, {
      method: "GET",
      headers: {
        apikey: API_KEY,
        Accept: "application/json",
      },
    });

    if (!response.ok) {
      const errorText = await response.text();
      console.error("❌ Direct fetch failed:", errorText);
      throw new Error(`Failed to fetch projects (${response.status}): ${errorText}`);
    }

    const jsonData = await response.json();
    console.log("✅ Direct fetch projects response:", jsonData);

    const data = jsonData.data || jsonData;
    if (Array.isArray(data)) {
      return data.map((p: any) => ({
        ...p,
        name: p.name || p.title || "Unnamed Project",
      }));
    }
    return [];
  };

  // 🔥 Direct fetch for tasks using /projects/items?projectid=
  const fetchTasksDirectly = async (projectId: number): Promise<Task[]> => {
    const API_KEY = localStorage.getItem("proworkflow-api-key") || process.env.PW_API_KEY || "";
    if (!API_KEY) {
      throw new Error("API key is not set.");
    }

    const url = `https://api.proworkflow.com/api/v4/projects/items?projectid=${projectId}`;
    console.log(`🔍 Direct fetch tasks for project ${projectId} from:`, url);

    const response = await fetch(url, {
      method: "GET",
      headers: {
        apikey: API_KEY,
        Accept: "application/json",
      },
    });

    if (!response.ok) {
      const errorText = await response.text();
      console.error("❌ Direct fetch tasks failed:", errorText);
      throw new Error(`Failed to fetch tasks (${response.status}): ${errorText}`);
    }

    const jsonData = await response.json();
    console.log(`✅ Direct fetch tasks response (project ${projectId}):`, jsonData);

    const data = jsonData.data || jsonData;
    if (Array.isArray(data)) {
      return data.map((item: any) => ({
        id: item.id,
        name: item.name || item.title || "Unnamed Task",
        projectid: item.projectid || projectId,
        assignee: item.assignee || item.assigned_to,
        description: item.description || "",
        priority: item.priority || "Medium",
        duedate: item.duedate || item.due_date || "",
        urgent: item.urgent || false,
        status: item.status || "",
        ...item,
      }));
    }
    return [];
  };

  // ----- Load data (projects and users) -----
  const loadData = async () => {
    setLoading(true);
    setError(null);
    try {
      const projectsData = await fetchProjectsDirectly();
      setProjects(projectsData);
      console.log(`📦 Loaded ${projectsData.length} projects via direct fetch`);

      const usersData = await proWorkflowApi.getUsers();
      setUsers(usersData);
      console.log(`👤 Loaded ${usersData.length} users via service`);
    } catch (err: any) {
      setError(err.message || "Failed to load data");
      console.error(err);
    } finally {
      setLoading(false);
    }
  };

  // ----- Load tasks using direct fetch -----
  const loadTasks = async (projectId: number) => {
    if (!projectId) {
      setTasks([]);
      return;
    }

    setLoadingTasks(true);
    setError(null);
    try {
      const tasksData = await fetchTasksDirectly(projectId);
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

  // ----- Initialize app -----
  useEffect(() => {
    if (isInitialized.current) return;
    isInitialized.current = true;

    const initializeApp = async () => {
      const defaultKey = (typeof process !== "undefined" && process.env.PW_API_KEY) || "";
      const savedKey = localStorage.getItem("proworkflow-api-key");
      const finalKey = savedKey && savedKey.trim() ? savedKey.trim() : defaultKey;

      if (finalKey) {
        setApiKey(finalKey);
        setIsApiKeySet(true);
        setApiKeyInput(finalKey);
        setError(null);
        await loadData();
      } else {
        setIsApiKeySet(false);
        setLoading(false);
      }
    };

    initializeApp();
  }, []);

  // ----- Project change handler -----
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

  // ----- Task selection handler -----
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

  // ----- Reset form -----
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

  // ----- Save (update) task using service -----
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

  // ----- Delete using service -----
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

  // ----- Cancel / Refresh / Toast helpers -----
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

  // ----- Filters and helpers -----
  const filteredTasks = tasks.filter(
    (task) =>
      task.name.toLowerCase().includes(searchTerm.toLowerCase()) ||
      (task.description && task.description.toLowerCase().includes(searchTerm.toLowerCase()))
  );

  const getUserName = (userId: number): string => {
    const user = users.find((u) => u.id === userId);
    return user ? user.name : "Unassigned";
  };

  // ----- Render -----
  if (!isApiKeySet && !loading) {
    return (
      <Box sx={{ p: 2 }}>
        <Paper sx={{ p: 3, textAlign: "center" }}>
          <Typography variant="h6" gutterBottom sx={{ fontWeight: 600, fontSize: "1rem" }}>
            🔑 ProWorkflow API Key Required
          </Typography>
          <Typography variant="body2" color="textSecondary" sx={{ mb: 2, fontSize: "0.8rem" }}>
            Please set your API key in the <strong>"New Task"</strong> tab first.
          </Typography>
          <Button
            variant="contained"
            size="small"
            onClick={() => {
              const tabs = document.querySelectorAll('[role="tab"]');
              if (tabs.length > 0) {
                (tabs[0] as HTMLElement).click();
              }
            }}
            sx={{ fontSize: "0.75rem" }}
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
        <Stack spacing={2}>
          <Skeleton variant="rectangular" height={36} animation="wave" />
          <Skeleton variant="rectangular" height={120} animation="wave" />
        </Stack>
      </Box>
    );
  }

  return (
    <Box sx={{ maxWidth: "100%" }}>
      {/* Compact Header */}
      <Stack direction="row" spacing={1} sx={{ alignItems: "center", mb: 1.5 }}>
        <Assignment color="primary" fontSize="small" />
        <Typography variant="subtitle1" sx={{ flex: 1, fontWeight: 600, fontSize: "0.9rem" }}>
          Edit Existing Task
        </Typography>
        <Tooltip title="Refresh data">
          <IconButton onClick={handleRefresh} size="small">
            <Refresh fontSize="small" />
          </IconButton>
        </Tooltip>
      </Stack>

      <Typography variant="body2" color="textSecondary" sx={{ mb: 1.5, fontSize: "0.7rem" }}>
        Select a project, then a task to view and edit its details.
      </Typography>

      {/* Alerts compact */}
      {error && (
        <Alert
          severity="error"
          sx={{ mb: 1, py: 0.5, fontSize: "0.7rem" }}
          onClose={() => setError(null)}
        >
          {error}
        </Alert>
      )}
      {success && (
        <Alert
          severity="success"
          sx={{ mb: 1, py: 0.5, fontSize: "0.7rem" }}
          onClose={() => setSuccess(false)}
        >
          <Stack direction="row" sx={{ alignItems: "center" }} spacing={0.5}>
            <CheckCircle fontSize="small" />
            <span>Task updated successfully! ✅</span>
          </Stack>
        </Alert>
      )}

      <Stack spacing={1.5}>
        {/* Project Selection */}
        <FormControl fullWidth size="small" sx={smallLabelStyles}>
          <InputLabel sx={{ fontSize: "0.7rem" }}>Select Project</InputLabel>
          <Select
            value={selectedProjectId}
            onChange={(e) => handleProjectChange(e.target.value as number)}
            label="Select Project"
            startAdornment={<Folder color="action" sx={{ mr: 1 }} fontSize="small" />}
          >
            <MenuItem value={0} sx={{ fontSize: "0.75rem" }}>
              Select a project...
            </MenuItem>
            {projects.map((project) => (
              <MenuItem key={project.id} value={project.id} sx={{ fontSize: "0.75rem" }}>
                {project.name}
              </MenuItem>
            ))}
          </Select>
        </FormControl>

        {/* Task Selection - Compact List */}
        {selectedProjectId > 0 && (
          <Paper variant="outlined" sx={{ p: 1.5, bgcolor: "#fafafa" }}>
            <Stack spacing={1}>
              <Stack
                direction="row"
                spacing={1}
                sx={{ alignItems: "center", pb: 0.5, borderBottom: "1px solid #e0e0e0" }}
              >
                <Typography variant="subtitle2" sx={{ fontWeight: 600, fontSize: "0.75rem" }}>
                  Available Tasks
                </Typography>
                <Chip
                  label={tasks.length}
                  size="small"
                  color="primary"
                  variant="filled"
                  sx={{ fontWeight: 600, height: 18, fontSize: "0.65rem" }}
                />
                <Box sx={{ flex: 1 }} />
                {loadingTasks && <CircularProgress size={14} />}
              </Stack>

              <TextField
                fullWidth
                size="small"
                placeholder="Search tasks..."
                value={searchTerm}
                onChange={(e) => setSearchTerm(e.target.value)}
                sx={{ "& .MuiOutlinedInput-root": { bgcolor: "white" }, ...smallLabelStyles }}
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
                <Box sx={{ maxHeight: "160px", overflow: "auto" }}>
                  {filteredTasks.map((task) => (
                    <Card
                      key={task.id}
                      variant="outlined"
                      sx={{
                        mb: 0.5,
                        cursor: "pointer",
                        borderColor: selectedTaskId === task.id ? "primary.main" : "#e0e0e0",
                        bgcolor: selectedTaskId === task.id ? "primary.light" : "white",
                        transition: "all 0.2s ease",
                        "&:hover": {
                          borderColor: "primary.main",
                          boxShadow: "0 1px 4px rgba(0,0,0,0.08)",
                        },
                      }}
                      onClick={() => handleTaskSelect(task.id)}
                    >
                      <CardContent sx={{ p: 0.75, "&:last-child": { pb: 0.75 } }}>
                        <Stack direction="row" spacing={0.75} sx={{ alignItems: "center" }}>
                          <Avatar
                            sx={{
                              width: 24,
                              height: 24,
                              bgcolor: getPriorityColor(task.priority || ""),
                              flexShrink: 0,
                            }}
                          >
                            {task.priority === "High" ? (
                              <PriorityHigh sx={{ fontSize: 14, color: "white" }} />
                            ) : (
                              <TaskAlt sx={{ fontSize: 14, color: "white" }} />
                            )}
                          </Avatar>

                          <Box sx={{ flex: 1, minWidth: 0 }}>
                            <Typography
                              variant="body2"
                              sx={{
                                fontWeight: selectedTaskId === task.id ? 600 : 500,
                                fontSize: "0.75rem",
                              }}
                            >
                              {task.name}
                            </Typography>
                            <Stack
                              direction="row"
                              spacing={0.5}
                              sx={{ mt: 0.25, alignItems: "center", flexWrap: "wrap" }}
                            >
                              <Chip
                                label={getStatusLabel(task.status)}
                                size="small"
                                sx={{
                                  bgcolor: getStatusColor(task.status || ""),
                                  color: "white",
                                  fontSize: "0.55rem",
                                  height: 16,
                                }}
                              />
                              {task.urgent && (
                                <Chip
                                  label="Urgent"
                                  size="small"
                                  color="error"
                                  sx={{ fontSize: "0.55rem", height: 16 }}
                                />
                              )}
                              {task.assignee && (
                                <Typography
                                  variant="caption"
                                  color="textSecondary"
                                  sx={{ fontSize: "0.6rem" }}
                                >
                                  Assigned to: {getUserName(task.assignee)}
                                </Typography>
                              )}
                              {task.duedate && (
                                <Typography
                                  variant="caption"
                                  color="textSecondary"
                                  sx={{ fontSize: "0.6rem" }}
                                >
                                  Due: {task.duedate}
                                </Typography>
                              )}
                            </Stack>
                          </Box>

                          {selectedTaskId === task.id && (
                            <CheckCircle color="primary" sx={{ fontSize: 14, flexShrink: 0 }} />
                          )}
                        </Stack>
                      </CardContent>
                    </Card>
                  ))}
                </Box>
              ) : (
                !loadingTasks && (
                  <Box sx={{ py: 1.5, textAlign: "center" }}>
                    <Info color="action" sx={{ fontSize: 24, opacity: 0.5, mb: 0.5 }} />
                    <Typography variant="body2" color="textSecondary" sx={{ fontSize: "0.75rem" }}>
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

        {/* Edit Form - Compact */}
        {isEditMode && originalTask && (
          <>
            <Divider />
            <Paper elevation={1} sx={{ p: 1.5 }}>
              <Stack spacing={1.5}>
                <Stack direction="row" spacing={1} sx={{ alignItems: "center", flexWrap: "wrap" }}>
                  <Badge color="primary" variant="dot">
                    <Typography variant="subtitle2" sx={{ fontWeight: 600, fontSize: "0.8rem" }}>
                      Editing: {originalTask.name}
                    </Typography>
                  </Badge>
                  <Box sx={{ flex: 1 }} />
                  <Chip
                    icon={<History fontSize="small" />}
                    label={`ID: #${originalTask.id}`}
                    size="small"
                    variant="outlined"
                    sx={{ height: 20, fontSize: "0.6rem" }}
                  />
                  <Button
                    variant="outlined"
                    color="error"
                    size="small"
                    startIcon={<Delete fontSize="small" />}
                    onClick={() => setShowDeleteDialog(true)}
                    sx={{ fontSize: "0.65rem" }}
                  >
                    Delete
                  </Button>
                </Stack>

                <Divider />

                {/* Status */}
                <FormControl fullWidth size="small" sx={smallLabelStyles}>
                  <InputLabel sx={{ fontSize: "0.7rem" }}>Status</InputLabel>
                  <Select value={status} onChange={(e) => setStatus(e.target.value)} label="Status">
                    <MenuItem value="" sx={{ fontSize: "0.75rem" }}>
                      Not Set
                    </MenuItem>
                    <MenuItem value="Not Started" sx={{ fontSize: "0.75rem" }}>
                      Not Started
                    </MenuItem>
                    <MenuItem value="In Progress" sx={{ fontSize: "0.75rem" }}>
                      In Progress
                    </MenuItem>
                    <MenuItem value="In Review" sx={{ fontSize: "0.75rem" }}>
                      In Review
                    </MenuItem>
                    <MenuItem value="Done" sx={{ fontSize: "0.75rem" }}>
                      Done
                    </MenuItem>
                    <MenuItem value="Blocked" sx={{ fontSize: "0.75rem" }}>
                      Blocked
                    </MenuItem>
                  </Select>
                </FormControl>

                {/* Task Name */}
                <TextField
                  required
                  fullWidth
                  size="small"
                  label="Task Name"
                  value={taskName}
                  onChange={(e) => setTaskName(e.target.value)}
                  placeholder="Enter task name..."
                  sx={smallLabelStyles}
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

                {/* Assignee */}
                <FormControl fullWidth size="small" sx={smallLabelStyles}>
                  <InputLabel sx={{ fontSize: "0.7rem" }}>Assignee</InputLabel>
                  <Select
                    value={assigneeId}
                    onChange={(e) => setAssigneeId(e.target.value as number)}
                    label="Assignee"
                    startAdornment={<Person color="action" sx={{ mr: 1 }} fontSize="small" />}
                  >
                    <MenuItem value={0} sx={{ fontSize: "0.75rem" }}>
                      Unassigned
                    </MenuItem>
                    {users.map((user) => (
                      <MenuItem key={user.id} value={user.id} sx={{ fontSize: "0.75rem" }}>
                        {user.name}
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
                <Grid container spacing={1.5}>
                  <Grid size={{ xs: 12, sm: 6 }}>
                    <FormControl fullWidth size="small" sx={smallLabelStyles}>
                      <InputLabel sx={{ fontSize: "0.7rem" }}>Priority</InputLabel>
                      <Select
                        value={priority}
                        onChange={(e) => setPriority(e.target.value as "Low" | "Medium" | "High")}
                        label="Priority"
                        startAdornment={
                          <PriorityHigh fontSize="small" color="action" sx={{ mr: 1 }} />
                        }
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
                  </Grid>
                  <Grid size={{ xs: 12, sm: 6 }}>
                    <TextField
                      fullWidth
                      size="small"
                      type="date"
                      label="Due Date"
                      value={dueDate}
                      onChange={(e) => setDueDate(e.target.value)}
                      sx={smallLabelStyles}
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
                  </Grid>
                </Grid>

                {/* Urgent Checkbox */}
                <FormControlLabel
                  control={
                    <Checkbox
                      checked={isUrgent}
                      onChange={(e) => setIsUrgent(e.target.checked)}
                      color="error"
                      size="small"
                    />
                  }
                  label={
                    <Stack direction="row" sx={{ alignItems: "center" }} spacing={0.5}>
                      <Warning color="error" fontSize="small" />
                      <Typography variant="body2" sx={{ fontSize: "0.75rem" }}>
                        Mark as Urgent
                      </Typography>
                    </Stack>
                  }
                />

                {/* Buttons */}
                <Stack direction="row" spacing={1.5} sx={{ justifyContent: "flex-end" }}>
                  <Button
                    variant="outlined"
                    size="small"
                    onClick={handleCancel}
                    disabled={saving}
                    startIcon={<Cancel fontSize="small" />}
                    sx={{ fontSize: "0.7rem" }}
                  >
                    Cancel
                  </Button>
                  <Button
                    variant="contained"
                    size="small"
                    onClick={handleSave}
                    disabled={saving || !taskName.trim()}
                    startIcon={saving ? <CircularProgress size={14} /> : <Save fontSize="small" />}
                    sx={{ fontSize: "0.7rem" }}
                  >
                    {saving ? "Saving..." : "Save Changes"}
                  </Button>
                </Stack>
              </Stack>
            </Paper>
          </>
        )}

        {selectedProjectId === 0 && (
          <Alert
            severity="info"
            icon={<Info fontSize="small" />}
            sx={{ py: 0.5, fontSize: "0.7rem" }}
          >
            Please select a project to view and edit tasks.
          </Alert>
        )}
      </Stack>

      {/* Delete Dialog */}
      <Dialog open={showDeleteDialog} onClose={() => setShowDeleteDialog(false)}>
        <DialogTitle sx={{ fontSize: "0.9rem" }}>Delete Task</DialogTitle>
        <DialogContent>
          <DialogContentText sx={{ fontSize: "0.8rem" }}>
            Are you sure you want to delete <strong>"{originalTask?.name}"</strong>? This action
            cannot be undone.
          </DialogContentText>
        </DialogContent>
        <DialogActions>
          <Button onClick={() => setShowDeleteDialog(false)} sx={{ fontSize: "0.75rem" }}>
            Cancel
          </Button>
          <Button
            onClick={handleDelete}
            color="error"
            variant="contained"
            disabled={deleting}
            startIcon={deleting ? <CircularProgress size={14} /> : <Delete />}
            sx={{ fontSize: "0.75rem" }}
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

export default EditTask;
