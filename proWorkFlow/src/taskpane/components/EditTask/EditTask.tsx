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
  Grid2,
  List,
  ListItem,
  ListItemText,
  ListItemIcon,
  Avatar,
  Badge,
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
  Pending,
  Block,
  Warning,
  Info,
  Assignment,
} from "@mui/icons-material";
import { editTaskService, mockTasks, mockUsers, mockProjects } from "../../services/editTask";
import { Project, User, Task } from "../../services/proworkflow";

// Helper function to get status color
const getStatusColor = (status: string): string => {
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

// Helper function to get priority color
const getPriorityColor = (priority: string): string => {
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
  const [loading, setLoading] = useState(false);
  const [loadingTasks, setLoadingTasks] = useState(false);
  const [saving, setSaving] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [success, setSuccess] = useState(false);
  const [isEditMode, setIsEditMode] = useState(false);
  const [searchTerm, setSearchTerm] = useState("");

  // Load initial data
  useEffect(() => {
    loadData();
  }, []);

  const loadData = async () => {
    setLoading(true);
    setError(null);
    try {
      const [projectsData, usersData] = await Promise.all([
        editTaskService.getProjects(),
        editTaskService.getUsers(),
      ]);
      setProjects(projectsData);
      setUsers(usersData);
    } catch (err) {
      setError("Failed to load data");
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
      const tasksData = await editTaskService.getTasks(projectId);
      setTasks(tasksData);

      // Reset selection
      setSelectedTaskId(0);
      setIsEditMode(false);
      resetForm();

      if (tasksData.length === 0) {
        setError("No tasks found in this project");
      }
    } catch (err) {
      setError("Failed to load tasks");
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

      const result = await editTaskService.updateTask(selectedTaskId, updatedData);
      console.log("Task updated:", result);

      // Update local tasks list
      const updatedTasks = tasks.map((task) =>
        task.id === selectedTaskId ? { ...task, ...updatedData } : task
      );
      setTasks(updatedTasks);

      // Update selected task
      setOriginalTask(result);

      setSuccess(true);
      setTimeout(() => setSuccess(false), 5000);
    } catch (err: any) {
      setError(err.message || "Failed to update task");
      console.error(err);
    } finally {
      setSaving(false);
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
  };

  // Filter tasks by search term
  const filteredTasks = tasks.filter(
    (task) =>
      task.name.toLowerCase().includes(searchTerm.toLowerCase()) ||
      (task.description && task.description.toLowerCase().includes(searchTerm.toLowerCase()))
  );

  // Get user name by id
  const getUserName = (userId: number): string => {
    const user = users.find((u) => u.id === userId);
    return user ? user.name : "Unassigned";
  };

  if (loading) {
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
        <Alert
          severity="error"
          sx={{ mb: 2 }}
          onClose={() => setError(null)}
          action={
            error.includes("No tasks") && selectedProjectId ? (
              <Button color="inherit" size="small" onClick={handleRefresh}>
                Refresh
              </Button>
            ) : undefined
          }
        >
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

        {/* Task Selection - Only show when project selected */}
        {selectedProjectId > 0 && (
          <>
            <Paper variant="outlined" sx={{ p: 2 }}>
              <Stack spacing={2}>
                <Stack direction="row" spacing={2} sx={{ alignItems: "center" }}>
                  <Typography variant="subtitle2">Available Tasks ({tasks.length})</Typography>
                  <Box sx={{ flex: 1 }} />
                  {loadingTasks && <CircularProgress size={20} />}
                </Stack>

                {/* Search */}
                <TextField
                  fullWidth
                  size="small"
                  placeholder="Search tasks..."
                  value={searchTerm}
                  onChange={(e) => setSearchTerm(e.target.value)}
                  slotProps={{
                    input: {
                      startAdornment: (
                        <InputAdornment position="start">
                          <Search fontSize="small" />
                        </InputAdornment>
                      ),
                    },
                  }}
                />

                {/* Task List */}
                {!loadingTasks && filteredTasks.length > 0 ? (
                  <List sx={{ maxHeight: "200px", overflow: "auto", bgcolor: "background.paper" }}>
                    {filteredTasks.map((task) => (
                      <ListItem
                        key={task.id}
                        component="div"
                        selected={selectedTaskId === task.id}
                        onClick={() => handleTaskSelect(task.id)}
                        sx={{
                          borderRadius: 1,
                          mb: 0.5,
                          cursor: "pointer",
                          "&:hover": {
                            bgcolor: "action.hover",
                          },
                          "&.Mui-selected": {
                            bgcolor: "primary.light",
                            "&:hover": {
                              bgcolor: "primary.light",
                            },
                          },
                        }}
                      >
                        <ListItemIcon>
                          <Avatar
                            sx={{
                              width: 32,
                              height: 32,
                              bgcolor: getPriorityColor(task.priority || ""),
                            }}
                          >
                            {task.priority === "High" ? (
                              <PriorityHigh fontSize="small" />
                            ) : (
                              <TaskAlt fontSize="small" />
                            )}
                          </Avatar>
                        </ListItemIcon>
                        <ListItemText
                          primary={task.name}
                          secondary={
                            <Stack
                              direction="row"
                              spacing={1}
                              sx={{ alignItems: "center", mt: 0.5 }}
                            >
                              <Chip
                                label={task.status || "Not Set"}
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
                            </Stack>
                          }
                        />
                        {task.duedate && (
                          <Typography variant="caption" color="textSecondary">
                            Due: {task.duedate}
                          </Typography>
                        )}
                      </ListItem>
                    ))}
                  </List>
                ) : (
                  !loadingTasks && (
                    <Alert severity="info" icon={<Info />}>
                      {searchTerm
                        ? "No tasks match your search"
                        : "No tasks available in this project"}
                    </Alert>
                  )
                )}
              </Stack>
            </Paper>
          </>
        )}

        {/* Edit Form - Only show when task selected */}
        {isEditMode && originalTask && (
          <>
            <Divider />

            <Paper elevation={2} sx={{ p: 3 }}>
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

                {/* Priority & Due Date - Using Grid2 */}
                <Grid2 container spacing={2}>
                  <Grid2 size={{ xs: 12, sm: 6 }}>
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
                  </Grid2>
                  <Grid2 size={{ xs: 12, sm: 6 }}>
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
                  </Grid2>
                </Grid2>

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

                {/* Change Summary */}
                {originalTask && (
                  <Paper variant="outlined" sx={{ p: 2, bgcolor: "#f8f9fa" }}>
                    <Typography
                      variant="caption"
                      color="textSecondary"
                      gutterBottom
                      sx={{ display: "block" }}
                    >
                      Changes will be applied to:
                    </Typography>
                    <Stack direction="row" spacing={2} sx={{ flexWrap: "wrap" }}>
                      <Chip
                        label={`Original: ${originalTask.name}`}
                        size="small"
                        variant="outlined"
                      />
                      {originalTask.status && (
                        <Chip
                          label={`Status: ${originalTask.status}`}
                          size="small"
                          variant="outlined"
                        />
                      )}
                      {originalTask.priority && (
                        <Chip
                          label={`Priority: ${originalTask.priority}`}
                          size="small"
                          variant="outlined"
                        />
                      )}
                    </Stack>
                  </Paper>
                )}

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
    </Box>
  );
};

export default EditTask;
