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

// Helper to convert Base64 to Blob
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
  const [contacts, setContacts] = useState<User[]>([]);
  const [loadingProjects, setLoadingProjects] = useState(false);
  const [loadingContacts, setLoadingContacts] = useState(false);

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

  // ----- Direct fetch for contacts -----
  const fetchContactsDirectly = async (): Promise<User[]> => {
    const API_KEY = localStorage.getItem("proworkflow-api-key") || process.env.PW_API_KEY || "";
    if (!API_KEY) {
      throw new Error("API key is not set.");
    }

    const url = "https://api.proworkflow.com/api/v4/contacts";
    console.log("🔍 Direct fetch contacts from:", url);

    const response = await fetch(url, {
      method: "GET",
      headers: {
        apikey: API_KEY,
        Accept: "application/json",
      },
    });

    if (!response.ok) {
      const errorText = await response.text();
      console.error("❌ Direct fetch contacts failed:", errorText);
      throw new Error(`Failed to fetch contacts (${response.status}): ${errorText}`);
    }

    const jsonData = await response.json();
    console.log("✅ Direct fetch contacts response:", jsonData);

    const data = jsonData.data || jsonData;
    if (Array.isArray(data)) {
      return data.map((c: any) => {
        let displayName = "Unnamed Contact";
        if (c.firstname && c.lastname) {
          displayName = `${c.firstname} ${c.lastname}`;
        } else if (c.firstname) {
          displayName = c.firstname;
        } else if (c.lastname) {
          displayName = c.lastname;
        } else if (c.companyname) {
          displayName = c.companyname;
        } else if (c.fullname) {
          displayName = c.fullname;
        }

        return {
          id: c.id,
          name: displayName,
          email: c.email || "",
          firstname: c.firstname || "",
          lastname: c.lastname || "",
          companyname: c.companyname || "",
          ...c,
        };
      });
    }
    return [];
  };

  // ----- Load Outlook data -----
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

  // ----- Load draft -----
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

  // ----- Load all data (projects via service, contacts via direct fetch) -----
  const loadAllData = useCallback(async () => {
    setLoadingData(true);
    setError(null);
    try {
      setLoadingProjects(true);
      setLoadingContacts(true);

      const projectsData = await proWorkflowApi.getProjects();
      setProjects(projectsData);
      setLoadingProjects(false);

      const contactsData = await fetchContactsDirectly();
      setContacts(contactsData);
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

  // Initialize app
  useEffect(() => {
    if (isInitialized.current) return;
    isInitialized.current = true;

    const initializeApp = async () => {
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

  // Auto-save draft
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

  // Direct create task using /projects/items
  const createTaskDirectly = async (taskData: any): Promise<any> => {
    const API_KEY = localStorage.getItem("proworkflow-api-key") || process.env.PW_API_KEY || "";
    if (!API_KEY) {
      throw new Error("API key is not set.");
    }

    const url = "https://api.proworkflow.com/api/v4/projects/items";
    console.log("🔍 Creating task via:", url);
    console.log("📦 Payload:", taskData);

    const response = await fetch(url, {
      method: "POST",
      headers: {
        apikey: API_KEY,
        "Content-Type": "application/json",
        Accept: "application/json",
      },
      body: JSON.stringify(taskData),
    });

    if (!response.ok) {
      const errorText = await response.text();
      console.error("❌ Create task failed:", errorText);
      throw new Error(`Failed to create task (${response.status}): ${errorText}`);
    }

    const jsonData = await response.json();
    console.log("✅ Task created response:", jsonData);
    return jsonData;
  };

  // ----- Submit handler with attachment upload -----
  const handleSubmit = async (e?: React.FormEvent | React.SyntheticEvent) => {
    if (e && e.preventDefault) {
      e.preventDefault();
    }

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
      const taskData = {
        projectid: projectId,
        name: taskName.trim(),
        description: description.trim() || undefined,
        assignee: assigneeId || undefined,
        priority: priority,
        duedate: dueDate || undefined,
        urgent: isUrgent,
      };

      const result = await createTaskDirectly(taskData);
      const newTaskId = result.id || result?.data?.id;
      if (!newTaskId) {
        throw new Error("Task created but no ID returned");
      }

      // Upload attachments if any
      if (includeAttachments && outlookData?.attachments && outlookData.attachments.length > 0) {
        console.log(`📎 Uploading ${outlookData.attachments.length} attachments...`);

        const uploadPromises = outlookData.attachments.map(async (att: any) => {
          let fileToUpload: File | Blob;
          let fileName = att.name || "attachment.bin";

          if (att instanceof File || att instanceof Blob) {
            fileToUpload = att;
          } else if (att.content && typeof att.content === "string") {
            const mimeType = att.mimeType || "application/octet-stream";
            fileToUpload = base64ToBlob(att.content, mimeType);
            fileName = att.name || "attachment.bin";
          } else if (typeof att === "string" && att.length > 0) {
            let mimeType = "application/octet-stream";
            let base64Data = att;
            if (att.startsWith("data:")) {
              const matches = att.match(/^data:([^;]+);/);
              if (matches && matches[1]) {
                mimeType = matches[1];
              }
              const parts = att.split(",");
              if (parts.length === 2) {
                base64Data = parts[1];
              } else {
                base64Data = att;
              }
            }
            fileToUpload = base64ToBlob(base64Data, mimeType);
          } else {
            console.warn("Skipping attachment with unknown format:", att);
            return null;
          }

          return (proWorkflowApi as any).uploadAttachment
            ? (proWorkflowApi as any).uploadAttachment(newTaskId, fileToUpload, fileName)
            : Promise.reject(new Error("uploadAttachment not implemented"));
        });

        const uploadResults = await Promise.all(
          uploadPromises.map((p) => p.catch((err) => ({ status: "rejected", reason: err })))
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
      setTimeout(() => setSuccess(false), 5000);
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

  // ----- Render -----
  if (!isApiKeySet && !loadingData) {
    return (
      <Box sx={{ p: 2 }}>
        <Paper sx={{ p: 3, textAlign: "center" }}>
          <Typography variant="h6" gutterBottom sx={{ fontWeight: 600 }}>
            🔑 ProWorkflow API Key
          </Typography>
          <Typography variant="body2" color="textSecondary" sx={{ mb: 2 }}>
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
              onKeyDown={(e) => {
                if (e.key === "Enter") {
                  handleApiKeySubmit();
                }
              }}
              sx={{ mb: 1.5 }}
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
            >
              {loading ? "Verifying..." : "Verify & Connect"}
            </Button>
            {error && (
              <Alert severity="error" sx={{ mt: 2 }}>
                {error}
              </Alert>
            )}
          </Box>
          <Paper variant="outlined" sx={{ p: 1.5, mt: 2, bgcolor: "#f8f9fa", textAlign: "left" }}>
            <Typography variant="caption" color="textSecondary" sx={{ fontWeight: 600 }}>
              💡 How to get your API key:
            </Typography>
            <Typography variant="body2" color="textSecondary" sx={{ mt: 0.5 }}>
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
        <Stack spacing={2}>
          <Skeleton variant="rectangular" height={40} animation="wave" />
          <Skeleton variant="rectangular" height={40} animation="wave" />
          <Skeleton variant="rectangular" height={40} animation="wave" />
          <Skeleton variant="rectangular" height={80} animation="wave" />
          <Skeleton variant="rectangular" height={40} animation="wave" />
          <Skeleton variant="rectangular" height={60} animation="wave" />
          <Skeleton variant="rectangular" height={36} animation="wave" />
        </Stack>
      </Box>
    );
  }

  return (
    <Box
      component="form"
      onSubmit={(e: React.FormEvent) => handleSubmit(e)}
      sx={{ maxWidth: "100%", p: 0 }}
    >
      {/* Outlook info header (compact) */}
      {outlookData && (
        <Paper
          elevation={0}
          sx={{
            p: 1,
            mb: 2,
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
            <Typography variant="caption" color="textSecondary">
              From:
            </Typography>
            <Chip
              label={outlookData.sender}
              size="small"
              variant="outlined"
              color="primary"
              sx={{ height: 20, fontSize: "0.7rem" }}
            />
            <Typography variant="caption" color="textSecondary" sx={{ ml: 0.5 }}>
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
                sx={{ mb: 0.5, display: "block" }}
              >
                Email Preview:
              </Typography>
              <Paper
                variant="outlined"
                sx={{
                  p: 0.5,
                  bgcolor: "white",
                  maxHeight: "100px",
                  overflow: "auto",
                  fontSize: "0.75rem",
                  whiteSpace: "pre-wrap",
                  wordBreak: "break-word",
                }}
              >
                {cleanEmailBody(outlookData.body || "No content") || "No content"}
              </Paper>
              {outlookData.attachments && outlookData.attachments.length > 0 && (
                <Box sx={{ mt: 0.5 }}>
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
        <Alert severity="info" sx={{ mb: 1.5, py: 0.5, fontSize: "0.75rem" }}>
          No email selected. Please select an email in Outlook to auto-fill task details.
          <Button size="small" onClick={handleRefreshOutlook} sx={{ ml: 1 }}>
            Refresh
          </Button>
        </Alert>
      )}

      {loadingOutlook && (
        <Box sx={{ display: "flex", alignItems: "center", mb: 1.5 }}>
          <CircularProgress size={16} sx={{ mr: 1 }} />
          <Typography variant="caption" color="textSecondary">
            Loading email data...
          </Typography>
        </Box>
      )}

      {error && (
        <Alert
          severity="error"
          sx={{ mb: 1.5, py: 0.5, fontSize: "0.75rem" }}
          onClose={() => setError(null)}
        >
          {error}
        </Alert>
      )}
      {success && (
        <Alert
          severity="success"
          sx={{ mb: 1.5, py: 0.5, fontSize: "0.75rem" }}
          icon={<CheckCircle />}
        >
          Task created successfully! 🎉
        </Alert>
      )}

      {draft && (
        <Alert severity="info" sx={{ mb: 1.5, py: 0.5, fontSize: "0.75rem" }} icon={<Warning />}>
          You have a saved draft. Continue editing or click Reset to clear.
        </Alert>
      )}

      <Stack spacing={1.5}>
        <TextField
          required
          size="small"
          fullWidth
          label="Task Name"
          value={taskName}
          onChange={(e) => setTaskName(e.target.value)}
          placeholder="Enter task name..."
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

        <FormControl fullWidth required size="small">
          <InputLabel>Project</InputLabel>
          <Select
            value={projectId}
            onChange={(e) => setProjectId(e.target.value as number)}
            label="Project"
            disabled={loadingProjects}
            startAdornment={
              loadingProjects ? (
                <CircularProgress size={16} sx={{ ml: 1 }} />
              ) : (
                <Folder color="action" sx={{ ml: 1 }} fontSize="small" />
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

        <FormControl fullWidth size="small">
          <InputLabel>Assignee</InputLabel>
          <Select
            value={assigneeId}
            onChange={(e) => setAssigneeId(e.target.value as number)}
            label="Assignee"
            disabled={loadingContacts}
            startAdornment={
              loadingContacts ? (
                <CircularProgress size={16} sx={{ ml: 1 }} />
              ) : (
                <Person color="action" sx={{ ml: 1 }} fontSize="small" />
              )
            }
          >
            <MenuItem value={0}>Unassigned</MenuItem>
            {contacts.map((contact) => (
              <MenuItem key={contact.id} value={contact.id}>
                {contact.name}
              </MenuItem>
            ))}
          </Select>
        </FormControl>

        <TextField
          fullWidth
          size="small"
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
                  <Description fontSize="small" color="action" sx={{ mr: 0.5 }} />
                </InputAdornment>
              ),
            },
          }}
        />

        <Stack direction="row" spacing={1.5}>
          <FormControl fullWidth size="small">
            <InputLabel>Priority</InputLabel>
            <Select
              value={priority}
              onChange={(e) => setPriority(e.target.value as "Low" | "Medium" | "High")}
              label="Priority"
              startAdornment={<PriorityHigh fontSize="small" color="action" sx={{ ml: 1 }} />}
            >
              <MenuItem value="Low">Low</MenuItem>
              <MenuItem value="Medium">Medium</MenuItem>
              <MenuItem value="High">High</MenuItem>
            </Select>
          </FormControl>

          <TextField
            fullWidth
            size="small"
            type="date"
            label="Due Date"
            value={dueDate}
            onChange={(e) => setDueDate(e.target.value)}
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

        <Paper variant="outlined" sx={{ p: 1 }}>
          <Stack spacing={1}>
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
                <Typography variant="body2" sx={{ fontSize: "0.8rem" }}>
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
                  size="small"
                />
              }
              label={
                <Stack direction="row" spacing={0.5} sx={{ alignItems: "center" }}>
                  <AttachFile fontSize="small" />
                  <Typography variant="body2" sx={{ fontSize: "0.8rem" }}>
                    Include email attachments{getAttachmentsLabel()}
                  </Typography>
                </Stack>
              }
            />
          </Stack>
        </Paper>

        <Typography
          variant="caption"
          color="textSecondary"
          sx={{ textAlign: "center", fontSize: "0.7rem" }}
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
          >
            Reset
          </Button>
          <Button
            type="button"
            variant="contained"
            size="small"
            disabled={loading || !taskName.trim() || !projectId}
            startIcon={loading ? <CircularProgress size={16} /> : <Send fontSize="small" />}
            onClick={handleSubmitClick}
          >
            {loading ? "Creating..." : "Create Task"}
          </Button>
        </Stack>
      </Stack>

      <Dialog open={showConfirmDialog} onClose={() => setShowConfirmDialog(false)}>
        <DialogTitle sx={{ fontSize: "1rem" }}>Confirm Action</DialogTitle>
        <DialogContent>
          <DialogContentText sx={{ fontSize: "0.9rem" }}>
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
          sx={{ fontSize: "0.8rem" }}
        >
          {toastMessage}
        </Alert>
      </Snackbar>
    </Box>
  );
};

export default CreateTask;
