// ProWorkflow API Service - Real API Implementation

const API_BASE_URL = 'https://api.proworkflow.com/api/v4';
let API_KEY = '';

export const setApiKey = (key: string) => {
  API_KEY = key;
};

export const getApiKey = (): string => {
  return API_KEY;
};

// API Headers
const getHeaders = () => ({
  'apikey': API_KEY,
  'Accept': 'application/json',
  'Content-Type': 'application/json',
});

// Types
export interface Project {
  id: number;
  name: string;
  description?: string;
  status?: string;
  companyid?: number;
}

export interface User {
  id: number;
  name: string;
  email?: string;
  username?: string;
  first_name?: string;
  last_name?: string;
}

export interface Contact {
  id: number;
  first_name: string;
  last_name: string;
  email?: string;
  username?: string;
  company_name?: string;
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
  estimated_hours?: number;
  billed_hours?: number;
  completed?: boolean;
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
  estimated_hours?: number;
  billed_hours?: number;
  completed?: boolean;
  created_date?: string;
  modified_date?: string;
}

// API Methods
export const proWorkflowApi = {
  // Test API Key
  testApiKey: async (apiKey: string): Promise<boolean> => {
    try {
      const response = await fetch(`${API_BASE_URL}/contacts?pagesize=1&pagenumber=1`, {
        method: 'GET',
        headers: {
          'apikey': apiKey,
          'Accept': 'application/json',
        },
      });
      return response.ok;
    } catch (error) {
      console.error('API Key test failed:', error);
      return false;
    }
  },

  // Get all projects
  getProjects: async (): Promise<Project[]> => {
    try {
      const response = await fetch(`${API_BASE_URL}/projects`, {
        method: 'GET',
        headers: getHeaders(),
      });
      
      if (!response.ok) {
        const errorData = await response.text();
        throw new Error(`Failed to fetch projects: ${response.status} - ${errorData}`);
      }
      
      const data = await response.json();
      return data.projects || [];
    } catch (error) {
      console.error('Error fetching projects:', error);
      throw error;
    }
  },

  // Get all users/contacts
  getUsers: async (): Promise<User[]> => {
    try {
      const response = await fetch(`${API_BASE_URL}/contacts`, {
        method: 'GET',
        headers: getHeaders(),
      });
      
      if (!response.ok) {
        const errorData = await response.text();
        throw new Error(`Failed to fetch users: ${response.status} - ${errorData}`);
      }
      
      const data = await response.json();
      // Map contacts to users format
      const contacts = data.contacts || [];
      return contacts.map((contact: Contact) => ({
        id: contact.id,
        name: `${contact.first_name || ''} ${contact.last_name || ''}`.trim() || contact.username || 'Unknown',
        email: contact.email || '',
        username: contact.username || '',
        first_name: contact.first_name || '',
        last_name: contact.last_name || '',
      }));
    } catch (error) {
      console.error('Error fetching users:', error);
      throw error;
    }
  },

  // Create a new task
  createTask: async (taskData: CreateTaskRequest): Promise<Task> => {
    try {
      // Clean data - remove undefined values
      const cleanData: any = { ...taskData };
      Object.keys(cleanData).forEach(key => {
        if (cleanData[key] === undefined || cleanData[key] === null) {
          delete cleanData[key];
        }
      });

      // Convert priority to string if needed
      if (cleanData.priority) {
        cleanData.priority = cleanData.priority.toString();
      }

      // Convert urgent to 0/1 if API expects it
      if (cleanData.urgent !== undefined) {
        cleanData.urgent = cleanData.urgent ? 1 : 0;
      }

      const response = await fetch(`${API_BASE_URL}/tasks`, {
        method: 'POST',
        headers: getHeaders(),
        body: JSON.stringify(cleanData),
      });
      
      if (!response.ok) {
        const errorData = await response.text();
        let errorMessage = `Failed to create task: ${response.status}`;
        try {
          const errorJson = JSON.parse(errorData);
          errorMessage = errorJson.message || errorJson.error || errorMessage;
        } catch {
          errorMessage = errorData || errorMessage;
        }
        throw new Error(errorMessage);
      }
      
      const data = await response.json();
      return data.task || data;
    } catch (error) {
      console.error('Error creating task:', error);
      throw error;
    }
  },

  // Get tasks for a project
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
        const errorData = await response.text();
        throw new Error(`Failed to fetch tasks: ${response.status} - ${errorData}`);
      }
      
      const data = await response.json();
      return data.tasks || [];
    } catch (error) {
      console.error('Error fetching tasks:', error);
      throw error;
    }
  },

  // Get single task details
  getTask: async (taskId: number): Promise<Task | null> => {
    try {
      const response = await fetch(`${API_BASE_URL}/tasks/${taskId}`, {
        method: 'GET',
        headers: getHeaders(),
      });
      
      if (!response.ok) {
        throw new Error(`Failed to fetch task: ${response.status}`);
      }
      
      const data = await response.json();
      return data.task || null;
    } catch (error) {
      console.error('Error fetching task:', error);
      throw error;
    }
  },

  // Update task
  updateTask: async (taskId: number, taskData: Partial<Task>): Promise<Task> => {
    try {
      // Clean data - remove undefined values
      const cleanData: any = { ...taskData };
      Object.keys(cleanData).forEach(key => {
        if (cleanData[key] === undefined || cleanData[key] === null) {
          delete cleanData[key];
        }
      });

      // Convert priority to string if needed
      if (cleanData.priority) {
        cleanData.priority = cleanData.priority.toString();
      }

      // Convert urgent to 0/1 if API expects it
      if (cleanData.urgent !== undefined) {
        cleanData.urgent = cleanData.urgent ? 1 : 0;
      }

      const response = await fetch(`${API_BASE_URL}/tasks/${taskId}`, {
        method: 'PUT',
        headers: getHeaders(),
        body: JSON.stringify(cleanData),
      });
      
      if (!response.ok) {
        const errorData = await response.text();
        let errorMessage = `Failed to update task: ${response.status}`;
        try {
          const errorJson = JSON.parse(errorData);
          errorMessage = errorJson.message || errorJson.error || errorMessage;
        } catch {
          errorMessage = errorData || errorMessage;
        }
        throw new Error(errorMessage);
      }
      
      const data = await response.json();
      return data.task || data;
    } catch (error) {
      console.error('Error updating task:', error);
      throw error;
    }
  },

  // Delete task
  deleteTask: async (taskId: number): Promise<boolean> => {
    try {
      const response = await fetch(`${API_BASE_URL}/tasks/${taskId}`, {
        method: 'DELETE',
        headers: getHeaders(),
      });
      
      if (!response.ok) {
        throw new Error(`Failed to delete task: ${response.status}`);
      }
      
      return true;
    } catch (error) {
      console.error('Error deleting task:', error);
      throw error;
    }
  },
};