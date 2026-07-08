// ProWorkflow API Service
const API_BASE_URL = 'https://api.proworkflow.net';
let API_KEY = 'your-api-key-here'; // Will be set from settings

export const setApiKey = (key: string) => {
  API_KEY = key;
};

// API Headers
const getHeaders = () => ({
  'Content-Type': 'application/json',
  'x-api-key': API_KEY,
});

// Types
export interface Project {
  id: number;
  name: string;
  description?: string;
  status?: string;
}

export interface User {
  id: number;
  name: string;
  email?: string;
  username?: string;
}

export interface CreateTaskRequest {
  projectid: number;
  name: string;
  description?: string;
  assignee?: number;
  priority?: 'Low' | 'Medium' | 'High';
  duedate?: string; // YYYY-MM-DD
  urgent?: boolean;
  attachments?: File[];
}

export interface Task {
  id: number;
  name: string;
  description?: string;
  projectid: number;
  assignee?: number;
  priority?: string;
  duedate?: string;
  urgent?: boolean;
  status?: string;
}

// API Methods
export const proWorkflowApi = {
  // Get all projects
  getProjects: async (): Promise<Project[]> => {
    try {
      const response = await fetch(`${API_BASE_URL}/projects`, {
        method: 'GET',
        headers: getHeaders(),
      });
      
      if (!response.ok) {
        throw new Error(`Failed to fetch projects: ${response.status}`);
      }
      
      const data = await response.json();
      return data.projects || [];
    } catch (error) {
      console.error('Error fetching projects:', error);
      throw error;
    }
  },

  // Get all users
  getUsers: async (): Promise<User[]> => {
    try {
      const response = await fetch(`${API_BASE_URL}/users`, {
        method: 'GET',
        headers: getHeaders(),
      });
      
      if (!response.ok) {
        throw new Error(`Failed to fetch users: ${response.status}`);
      }
      
      const data = await response.json();
      return data.users || [];
    } catch (error) {
      console.error('Error fetching users:', error);
      throw error;
    }
  },

  // Create a new task
  createTask: async (taskData: CreateTaskRequest): Promise<Task> => {
    try {
      // Remove undefined values
      const cleanData: any = { ...taskData };
      Object.keys(cleanData).forEach(key => {
        if (cleanData[key] === undefined) {
          delete cleanData[key];
        }
      });

      // If attachments are included, handle them separately
      const { attachments, ...taskWithoutAttachments } = cleanData;

      const response = await fetch(`${API_BASE_URL}/tasks`, {
        method: 'POST',
        headers: getHeaders(),
        body: JSON.stringify(taskWithoutAttachments),
      });
      
      if (!response.ok) {
        const errorData = await response.json();
        throw new Error(errorData.message || `Failed to create task: ${response.status}`);
      }
      
      const data = await response.json();
      return data.task || data;
    } catch (error) {
      console.error('Error creating task:', error);
      throw error;
    }
  },

  // Get tasks (for edit functionality)
  getTasks: async (projectId?: number): Promise<Task[]> => {
    try {
      const url = projectId 
        ? `${API_BASE_URL}/tasks?projectid=${projectId}`
        : `${API_BASE_URL}/tasks`;
      
      const response = await fetch(url, {
        method: 'GET',
        headers: getHeaders(),
      });
      
      if (!response.ok) {
        throw new Error(`Failed to fetch tasks: ${response.status}`);
      }
      
      const data = await response.json();
      return data.tasks || [];
    } catch (error) {
      console.error('Error fetching tasks:', error);
      throw error;
    }
  },
};