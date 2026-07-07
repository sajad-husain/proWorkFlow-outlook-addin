import axios from 'axios';
import { mockProjects, mockAssignees, mockTaskCreationResponse } from './mockData';

// TODO: Replace with actual API configuration when you have credentials
const API_CONFIG = {
  baseURL: 'https://api.proworkflow.net/v1',
  apiKey: 'YOUR_API_KEY',
  username: 'YOUR_USERNAME',
  // For Nexus v4 API, use bearer token instead
};

// Mock mode - set to false when ready for real API
let USE_MOCK = true;

const apiClient = axios.create({
  baseURL: API_CONFIG.baseURL,
  // For Classic API - Basic Auth
  auth: {
    username: API_CONFIG.username,
    password: API_CONFIG.apiKey,
  },
  headers: {
    'Content-Type': 'application/json',
  },
});

export const proWorkflowApi = {
  getProjects: async () => {
    if (USE_MOCK) {
      console.log('[MOCK] Fetching projects...');
      return mockProjects;
    }
    try {
      const response = await apiClient.get('/projects');
      return response.data;
    } catch (error) {
      console.error('Error fetching projects:', error);
      throw error;
    }
  },

  getAssignees: async () => {
    if (USE_MOCK) {
      console.log('[MOCK] Fetching assignees...');
      return mockAssignees;
    }
    try {
      const response = await apiClient.get('/contacts');
      return response.data;
    } catch (error) {
      console.error('Error fetching assignees:', error);
      throw error;
    }
  },

  createTask: async (taskData) => {
    if (USE_MOCK) {
      console.log('[MOCK] Creating task:', taskData);
      await new Promise(resolve => setTimeout(resolve, 1000));
      return mockTaskCreationResponse;
    }
    try {
      const response = await apiClient.post('/tasks', taskData);
      return response.data;
    } catch (error) {
      console.error('Error creating task:', error);
      throw error;
    }
  },
};

export const setApiMode = (useMock) => {
  USE_MOCK = useMock;
  console.log(`API mode set to ${useMock ? 'MOCK' : 'REAL'}`);
};