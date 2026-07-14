// Edit Task Service - Using Real API

import { Task, Project, User } from './proworkflow';
import { proWorkflowApi } from './proworkflow';

// Edit Task API Service
export const editTaskService = {
  // Get tasks for a project
  getTasks: async (projectId: number): Promise<Task[]> => {
    try {
      const tasks = await proWorkflowApi.getTasks(projectId);
      return tasks;
    } catch (error) {
      console.error('Error loading tasks:', error);
      throw error;
    }
  },

  // Get all projects
  getProjects: async (): Promise<Project[]> => {
    try {
      const projects = await proWorkflowApi.getProjects();
      return projects;
    } catch (error) {
      console.error('Error loading projects:', error);
      throw error;
    }
  },

  // Get all users
  getUsers: async (): Promise<User[]> => {
    try {
      const users = await proWorkflowApi.getUsers();
      return users;
    } catch (error) {
      console.error('Error loading users:', error);
      throw error;
    }
  },

  // Update task
  updateTask: async (taskId: number, taskData: Partial<Task>): Promise<Task> => {
    try {
      const updatedTask = await proWorkflowApi.updateTask(taskId, taskData);
      return updatedTask;
    } catch (error) {
      console.error('Error updating task:', error);
      throw error;
    }
  },

  // Get single task
  getTask: async (taskId: number): Promise<Task | null> => {
    try {
      const task = await proWorkflowApi.getTask(taskId);
      return task;
    } catch (error) {
      console.error('Error fetching task:', error);
      throw error;
    }
  },

  // Delete task
  deleteTask: async (taskId: number): Promise<boolean> => {
    try {
      return await proWorkflowApi.deleteTask(taskId);
    } catch (error) {
      console.error('Error deleting task:', error);
      throw error;
    }
  },
};