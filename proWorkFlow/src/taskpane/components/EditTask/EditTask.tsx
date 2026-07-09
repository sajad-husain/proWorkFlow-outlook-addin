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
} from "@mui/icons-material";
import { proWorkflowApi, Project, User, Task } from "../../services/proworkflow";

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

  // UI state
  const [loading, setLoading] = useState(false);
  const [loadingTasks, setLoadingTasks] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [success, setSuccess] = useState(false);
  const [isEditMode, setIsEditMode] = useState(false);

  // Load initial data
  useEffect(() => {
    loadProjectsAndUsers();
  }, []);

  const loadProjectsAndUsers = async () => {
    setLoading(true);
    try {
      const [projectsData, usersData] = await Promise.all([
        proWorkflowApi.getProjects(),
        proWorkflowApi.getUsers(),
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
    try {
      const tasksData = await proWorkflowApi.getTasks(projectId);
      setTasks(tasksData);
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
    if (projectId) {
      loadTasks(projectId);
    }
  };

  const handleTaskSelect = (taskId: number) => {
    setSelectedTaskId(taskId);
    setIsEditMode(true);

    const task = tasks.find((t) => t.id === taskId);
    if (task) {
      setTaskName(task.name || "");
      setAssigneeId(task.assignee || 0);
      setDescription(task.description || "");
      setPriority((task.priority as "Low" | "Medium" | "High") || "Medium");
      setDueDate(task.duedate || "");
      setIsUrgent(task.urgent || false);
      setStatus(task.status || "");
    }
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

    setLoading(true);
    setError(null);

    try {
      // TODO: Update task API call
      // await proWorkflowApi.updateTask(selectedTaskId, {
      //   name: taskName,
      //   assignee: assigneeId || undefined,
      //   description: description || undefined,
      //   priority: priority,
      //   duedate: dueDate || undefined,
      //   urgent: isUrgent,
      //   status: status || undefined,
      // });

      setSuccess(true);
      setTimeout(() => setSuccess(false), 5000);
    } catch (err: any) {
      setError(err.message || "Failed to update task");
    } finally {
      setLoading(false);
    }
  };

  const handleReset = () => {
    setSelectedTaskId(0);
    setIsEditMode(false);
    setTaskName("");
    setAssigneeId(0);
    setDescription("");
    setPriority("Medium");
    setDueDate("");
    setIsUrgent(false);
    setStatus("");
    setError(null);
    setSuccess(false);
  };

  return (
    <Box sx={{ maxWidth: "100%" }}>
      <Typography variant="h6" gutterBottom>
        Edit Existing Task
      </Typography>
      <Typography variant="body2" color="textSecondary" sx={{ mb: 3 }}>
        Select a project, then a task to edit its details.
      </Typography>

      {error && (
        <Alert severity="error" sx={{ mb: 2 }} onClose={() => setError(null)}>
          {error}
        </Alert>
      )}
      {success && (
        <Alert severity="success" sx={{ mb: 2 }}>
          Task updated successfully! ✅
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

        {/* Task Selection */}
        {selectedProjectId > 0 && (
          <FormControl fullWidth>
            <InputLabel>Select Task</InputLabel>
            <Select
              value={selectedTaskId}
              onChange={(e) => handleTaskSelect(e.target.value as number)}
              label="Select Task"
              disabled={loadingTasks}
              startAdornment={<Search color="action" sx={{ mr: 1 }} />}
            >
              <MenuItem value={0}>Select a task...</MenuItem>
              {tasks.map((task) => (
                <MenuItem key={task.id} value={task.id}>
                  {task.name}
                </MenuItem>
              ))}
            </Select>
          </FormControl>
        )}

        {loadingTasks && (
          <Box sx={{ display: "flex", justifyContent: "center" }}>
            <CircularProgress size={24} />
            <Typography variant="caption" sx={{ ml: 1 }}>
              Loading tasks...
            </Typography>
          </Box>
        )}

        {/* Edit Form */}
        {isEditMode && (
          <>
            <Divider />

            <Paper variant="outlined" sx={{ p: 3 }}>
              <Stack spacing={3}>
                {/* Status */}
                <TextField
                  fullWidth
                  label="Status"
                  value={status}
                  onChange={(e) => setStatus(e.target.value)}
                  placeholder="e.g., In Progress, Review, Done"
                  disabled
                  helperText="Status updates coming soon"
                />

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

                {/* Priority & Due Date */}
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

                {/* Urgent Checkbox */}
                <FormControlLabel
                  control={
                    <Checkbox
                      checked={isUrgent}
                      onChange={(e) => setIsUrgent(e.target.checked)}
                      color="error"
                    />
                  }
                  label="Mark as Urgent"
                />

                {/* Action Buttons */}
                <Stack direction="row" spacing={2} sx={{ justifyContent: "flex-end" }}>
                  <Button
                    variant="outlined"
                    onClick={handleReset}
                    disabled={loading}
                    startIcon={<Cancel />}
                  >
                    Cancel
                  </Button>
                  <Button
                    variant="contained"
                    onClick={handleSave}
                    disabled={loading || !taskName.trim()}
                    startIcon={loading ? <CircularProgress size={20} /> : <Save />}
                  >
                    {loading ? "Saving..." : "Save Changes"}
                  </Button>
                </Stack>
              </Stack>
            </Paper>
          </>
        )}

        {selectedProjectId > 0 && !isEditMode && !loadingTasks && tasks.length === 0 && (
          <Alert severity="info">
            No tasks found in this project. Switch to "New Task" tab to create one.
          </Alert>
        )}
      </Stack>
    </Box>
  );
};

export default EditTask;
