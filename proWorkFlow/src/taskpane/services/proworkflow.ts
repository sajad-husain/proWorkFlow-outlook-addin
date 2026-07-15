// src/services/proworkflow.ts

const BASE_URL = process.env.PW_BASE_URL || "https://api.proworkflow.com/api/v4/";

let API_KEY = process.env.PW_API_KEY || "";

export const setApiKey = (key: string) => {
  API_KEY = key.trim();
};

// Helper: API fetch with error handling
const apiFetch = async <T>(endpoint: string, options?: RequestInit): Promise<T> => {
  if (!API_KEY) {
    throw new Error("API key is not set.");
  }

  const url = `${BASE_URL}${endpoint}`;
  const response = await fetch(url, {
    ...options,
    headers: {
      "apikey": API_KEY,
      "Accept": "application/json",
      ...(options?.headers || {}),
    },
  });

  if (!response.ok) {
    const errorText = await response.text();
    throw new Error(`API Error (${response.status}): ${errorText || response.statusText}`);
  }

  // DELETE returns 204 No Content
  if (response.status === 204) {
    return {} as T;
  }

  return response.json();
};

// ---------- INTERFACES ----------
export interface Project {
  id: number;
  name?: string;
  title?: string;
  number?: string;
  startdate?: string;
  duedate?: string;
  companyid?: number;
  companyname?: string;
}

export interface User {
  id: number;
  name: string;
  email?: string;
}

export interface Task {
  data?: any;
  id: number;
  name: string;
  projectid: number;
  assignee?: number;
  description?: string;
  priority?: "Low" | "Medium" | "High";
  duedate?: string;
  urgent?: boolean;
  status?: string;
}

export interface CreateTaskRequest {
  name: string;
  projectid: number;
  assignee?: number;
  description?: string;
  priority?: "Low" | "Medium" | "High";
  duedate?: string;
  urgent?: boolean;
}

// ---------- API METHODS ----------
export const proWorkflowApi = {
  // Test API key
  testApiKey: async (key: string): Promise<boolean> => {
    try {
      const tempKey = API_KEY;
      setApiKey(key);
      await apiFetch<any>("projects?limit=1");
      setApiKey(tempKey);
      return true;
    } catch {
      return false;
    }
  },

  // Get all projects
  getProjects: async (): Promise<Project[]> => {
    const response = await apiFetch<any>("projects");
    const data = response.data || response;
    if (Array.isArray(data)) {
      return data.map((p: any) => ({
        ...p,
        name: p.name || p.title || "Unnamed Project",
      }));
    }
    return [];
  },

  // Get all users
  getUsers: async (): Promise<User[]> => {
    const response = await apiFetch<any>("users");
    const data = response.data || response;
    return Array.isArray(data) ? data : [];
  },

  // Get tasks for a project
  getTasks: async (projectId: number): Promise<Task[]> => {
    const response = await apiFetch<any>(`tasks?projectid=${projectId}`);
    const data = response.data || response;
    return Array.isArray(data) ? data : [];
  },

  // Get a single task by ID
  getTask: async (taskId: number): Promise<Task> => {
    const response = await apiFetch<any>(`tasks/${taskId}`);
    return response.data || response;
  },

  // Create task
  createTask: async (taskData: CreateTaskRequest): Promise<Task> => {
    const response = await apiFetch<any>("tasks", {
      method: "POST",
      body: JSON.stringify(taskData),
      headers: { "Content-Type": "application/json" },
    });
    return response.data || response;
  },

  // Update task
  updateTask: async (
    taskId: number,
    taskData: Partial<CreateTaskRequest> & { status?: string }
  ): Promise<Task> => {
    const response = await apiFetch<any>(`tasks/${taskId}`, {
      method: "PUT",
      body: JSON.stringify(taskData),
      headers: { "Content-Type": "application/json" },
    });
    return response.data || response;
  },

  // Delete task
  deleteTask: async (taskId: number): Promise<boolean> => {
    try {
      await apiFetch<void>(`tasks/${taskId}`, { method: "DELETE" });
      return true;
    } catch {
      return false;
    }
  },

  // ✅ Upload attachment to a task (added inside object)
  uploadAttachment: async (taskId: number, file: File | Blob, fileName: string): Promise<any> => {
    const formData = new FormData();
    formData.append("file", file, fileName);

    const response = await fetch(`${BASE_URL}tasks/${taskId}/attachments`, {
      method: "POST",
      headers: {
        apikey: API_KEY,
        // Content-Type automatically set by FormData
      },
      body: formData,
    });

    if (!response.ok) {
      const errorText = await response.text();
      throw new Error(`Attachment upload failed (${response.status}): ${errorText}`);
    }

    return response.json();
  },
};