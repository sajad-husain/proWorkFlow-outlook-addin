import * as React from "react";
import {
  makeStyles,
  tokens,
  Text,
  Input,
  Textarea,
  Button,
  Select,
  Checkbox,
  Card,
  CardHeader,
  CardPreview,
  Spinner,
  Field,
  Divider,
  Toast,
  ToastTitle,
  ToastBody,
  Toaster,
  useToastController,
} from "@fluentui/react-components";
import {
  Add24Regular,
  Calendar24Regular,
  Person24Regular,
  Flag24Regular,
  CheckmarkCircle24Regular,
} from "@fluentui/react-icons";

const useStyles = makeStyles({
  formContainer: {
    padding: "20px",
    flex: 1,
    overflowY: "auto",
  },
  card: {
    marginBottom: "16px",
    padding: "16px",
  },
  cardHeader: {
    fontSize: "14px",
    fontWeight: 600,
    color: tokens.colorNeutralForeground1,
    marginBottom: "12px",
  },
  field: {
    marginBottom: "16px",
    width: "100%",
  },
  fieldLabel: {
    fontSize: "13px",
    fontWeight: 500,
    color: tokens.colorNeutralForeground2,
    marginBottom: "4px",
    display: "block",
  },
  row: {
    display: "flex",
    gap: "12px",
    marginBottom: "12px",
    flexWrap: "wrap",
  },
  flex1: {
    flex: 1,
    minWidth: "120px",
  },
  checkboxGroup: {
    display: "flex",
    gap: "20px",
    alignItems: "center",
    padding: "8px 0",
  },
  submitButton: {
    width: "100%",
    padding: "12px",
    backgroundColor: "#e94560",
    color: "#ffffff",
    fontSize: "16px",
    fontWeight: 600,
    border: "none",
    borderRadius: "6px",
    cursor: "pointer",
    transition: "background-color 0.2s",
    "&:hover": {
      backgroundColor: "#c73652",
    },
    "&:disabled": {
      opacity: 0.6,
      cursor: "not-allowed",
    },
  },
  loadingOverlay: {
    display: "flex",
    justifyContent: "center",
    alignItems: "center",
    padding: "40px 0",
  },
  emailPreview: {
    backgroundColor: tokens.colorNeutralBackground2,
    padding: "12px",
    borderRadius: "6px",
    marginBottom: "16px",
    fontSize: "13px",
    color: tokens.colorNeutralForeground2,
    border: `1px solid ${tokens.colorNeutralStroke1}`,
  },
  emailSubject: {
    fontWeight: 600,
    color: tokens.colorNeutralForeground1,
    fontSize: "14px",
  },
});

interface CreateTaskFormProps {
  emailData: {
    subject: string;
    body: string;
    from: string;
    attachments?: any[];
  } | null;
  loading: boolean;
}

const CreateTaskForm: React.FC<CreateTaskFormProps> = ({ emailData, loading }) => {
  const styles = useStyles();
  const { dispatchToast } = useToastController();

  // Mock data - will be replaced with real API calls
  const [workspaces] = React.useState([
    { id: "1", name: "Discover Redbooth" },
    { id: "2", name: "Marketing Projects" },
    { id: "3", name: "Development Team" },
  ]);

  const [taskLists] = React.useState([
    { id: "1", name: "Get Started", workspaceId: "1" },
    { id: "2", name: "Active Tasks", workspaceId: "1" },
    { id: "3", name: "Backlog", workspaceId: "1" },
  ]);

  const [assignees] = React.useState([
    { id: "1", name: "Joshua Von Bischoffshausen" },
    { id: "2", name: "Sarah Chen" },
    { id: "3", name: "Mike Johnson" },
  ]);

  const [formData, setFormData] = React.useState({
    workspace: "",
    taskList: "",
    task: "",
    title: "",
    description: "",
    assignee: "",
    dueDate: "",
    isUrgent: false,
    addAttachments: false,
  });

  const [isSubmitting, setIsSubmitting] = React.useState(false);

  // Pre-fill title from email subject
  React.useEffect(() => {
    if (emailData?.subject) {
      setFormData((prev) => ({
        ...prev,
        title: emailData.subject,
        description: emailData.body || "",
      }));
    }
  }, [emailData]);

  const handleChange = (
    e: React.ChangeEvent<HTMLInputElement | HTMLTextAreaElement | HTMLSelectElement>
  ) => {
    const { name, value, type } = e.target;
    setFormData((prev) => ({
      ...prev,
      [name]: type === "checkbox" ? (e.target as HTMLInputElement).checked : value,
    }));
  };

  const handleCheckboxChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    const { name, checked } = e.target;
    setFormData((prev) => ({
      ...prev,
      [name]: checked,
    }));
  };

  const handleSubmit = async (e: React.FormEvent) => {
    e.preventDefault();
    setIsSubmitting(true);

    try {
      // TODO: Replace with actual ProWorkflow API call
      console.log("Creating task:", formData);

      // Simulate API call
      await new Promise((resolve) => setTimeout(resolve, 1500));

      dispatchToast(
        <Toast>
          <ToastTitle>Task Created Successfully!</ToastTitle>
          <ToastBody>Your task has been created in ProWorkflow.</ToastBody>
        </Toast>,
        { intent: "success", position: "top-end" }
      );

      // Reset form
      setFormData((prev) => ({
        ...prev,
        workspace: "",
        taskList: "",
        task: "",
        assignee: "",
        dueDate: "",
        isUrgent: false,
        addAttachments: false,
      }));
    } catch (error) {
      dispatchToast(
        <Toast>
          <ToastTitle>Error Creating Task</ToastTitle>
          <ToastBody>Please try again later.</ToastBody>
        </Toast>,
        { intent: "error", position: "top-end" }
      );
    } finally {
      setIsSubmitting(false);
    }
  };

  if (loading) {
    return (
      <div className={styles.loadingOverlay}>
        <Spinner label="Loading email..." />
      </div>
    );
  }

  return (
    <div className={styles.formContainer}>
      <Toaster />

      {/* Email Preview */}
      {emailData && (
        <div className={styles.emailPreview}>
          <Text className={styles.emailSubject}>{emailData.subject}</Text>
          {emailData.from && (
            <div>
              <Person24Regular
                style={{ display: "inline", marginRight: "4px", fontSize: "14px" }}
              />
              <Text style={{ fontSize: "12px", color: tokens.colorNeutralForeground3 }}>
                From: {emailData.from}
              </Text>
            </div>
          )}
        </div>
      )}

      <form onSubmit={handleSubmit}>
        {/* Workspace */}
        <div className={styles.field}>
          <label className={styles.fieldLabel}>Workspace</label>
          <Select
            name="workspace"
            value={formData.workspace}
            onChange={handleChange}
            placeholder="Select a workspace..."
            style={{ width: "100%" }}
          >
            <option value="">Select a workspace...</option>
            {workspaces.map((w) => (
              <option key={w.id} value={w.id}>
                {w.name}
              </option>
            ))}
          </Select>
        </div>

        {/* Task List */}
        <div className={styles.field}>
          <label className={styles.fieldLabel}>Task list</label>
          <Select
            name="taskList"
            value={formData.taskList}
            onChange={handleChange}
            placeholder="Select a task list..."
            style={{ width: "100%" }}
            disabled={!formData.workspace}
          >
            <option value="">Select a task list...</option>
            {taskLists
              .filter((tl) => tl.workspaceId === formData.workspace || !formData.workspace)
              .map((tl) => (
                <option key={tl.id} value={tl.id}>
                  {tl.name}
                </option>
              ))}
          </Select>
        </div>

        {/* Task (optional - for editing existing tasks) */}
        <div className={styles.field}>
          <label className={styles.fieldLabel}>Task</label>
          <Select
            name="task"
            value={formData.task}
            onChange={handleChange}
            placeholder="Select a task..."
            style={{ width: "100%" }}
          >
            <option value="">Select a task...</option>
            <option value="new">+ Create new task</option>
          </Select>
        </div>

        <Divider style={{ margin: "16px 0" }} />

        {/* Title */}
        <div className={styles.field}>
          <label className={styles.fieldLabel}>Title</label>
          <Input
            name="title"
            value={formData.title}
            onChange={handleChange}
            placeholder="Enter task title..."
            style={{ width: "100%" }}
            required
          />
        </div>

        {/* Description */}
        <div className={styles.field}>
          <label className={styles.fieldLabel}>Description</label>
          <Textarea
            name="description"
            value={formData.description}
            onChange={handleChange}
            placeholder="Enter task description..."
            rows={4}
            style={{ width: "100%" }}
          />
        </div>

        {/* Assignee & Due Date Row */}
        <div className={styles.row}>
          <div className={styles.flex1}>
            <label className={styles.fieldLabel}>Assignee</label>
            <Select
              name="assignee"
              value={formData.assignee}
              onChange={handleChange}
              style={{ width: "100%" }}
            >
              <option value="">Select assignee...</option>
              {assignees.map((a) => (
                <option key={a.id} value={a.id}>
                  {a.name}
                </option>
              ))}
            </Select>
          </div>
          <div className={styles.flex1}>
            <label className={styles.fieldLabel}>Due date</label>
            <Input
              type="date"
              name="dueDate"
              value={formData.dueDate}
              onChange={handleChange}
              style={{ width: "100%" }}
            />
          </div>
        </div>

        {/* Checkboxes */}
        <div className={styles.checkboxGroup}>
          <Checkbox
            name="isUrgent"
            label="Marks as Urgent"
            checked={formData.isUrgent}
            onChange={handleCheckboxChange}
          />
          <Checkbox
            name="addAttachments"
            label="Add attachments"
            checked={formData.addAttachments}
            onChange={handleCheckboxChange}
          />
        </div>

        <Divider style={{ margin: "16px 0" }} />

        {/* Submit Button */}
        <button type="submit" className={styles.submitButton} disabled={isSubmitting}>
          {isSubmitting ? (
            <span
              style={{
                display: "flex",
                alignItems: "center",
                justifyContent: "center",
                gap: "8px",
              }}
            >
              <Spinner size="tiny" /> Creating...
            </span>
          ) : (
            "Create task"
          )}
        </button>
      </form>
    </div>
  );
};

export default CreateTaskForm;
