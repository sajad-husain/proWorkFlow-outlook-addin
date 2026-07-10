import { Task, Project, User } from './proworkflow';

// Mock data for testing 
export const mockTasks: Task[] = [
  {
    id: 1,
    name: 'Review Q4 Financial Report',
    description: 'Need to review and approve the quarterly financial report before board meeting',
    projectid: 101,
    assignee: 201,
    priority: 'High',
    duedate: '2026-07-15',
    urgent: true,
    status: 'In Progress',
  },
  {
    id: 2,
    name: 'Design New Landing Page',
    description: 'Create mockups for the new product landing page with updated branding',
    projectid: 101,
    assignee: 202,
    priority: 'Medium',
    duedate: '2026-07-20',
    urgent: false,
    status: 'Not Started',
  },
  {
    id: 3,
    name: 'Update User Documentation',
    description: 'Update API documentation with new endpoints and examples',
    projectid: 102,
    assignee: 203,
    priority: 'Low',
    duedate: '2026-07-25',
    urgent: false,
    status: 'In Review',
  },
  {
    id: 4,
    name: 'Fix Login Bug',
    description: 'Users unable to login after recent update. Need immediate fix.',
    projectid: 102,
    assignee: 201,
    priority: 'High',
    duedate: '2026-07-10',
    urgent: true,
    status: 'Done',
  },
  {
    id: 5,
    name: 'Prepare Sales Presentation',
    description: 'Create presentation for upcoming client meeting about new features',
    projectid: 103,
    assignee: 204,
    priority: 'Medium',
    duedate: '2026-07-18',
    urgent: false,
    status: 'In Progress',
  },
];

// Mock projects for dropdown
export const mockProjects = [
  { id: 101, name: 'Product Development' },
  { id: 102, name: 'Bug Fixes' },
  { id: 103, name: 'Sales & Marketing' },
  { id: 104, name: 'Documentation' },
];

// Mock users for dropdown
export const mockUsers = [
  { id: 201, name: 'John Smith', email: 'john@example.com' },
  { id: 202, name: 'Sarah Johnson', email: 'sarah@example.com' },
  { id: 203, name: 'Mike Brown', email: 'mike@example.com' },
  { id: 204, name: 'Emily Davis', email: 'emily@example.com' },
];

// Edit Task API Service (will use real API when available)
export const editTaskService = {
  // Get tasks for a project
  getTasks: async (projectId: number): Promise<Task[]> => {
    // Simulate API call
    await new Promise(resolve => setTimeout(resolve, 500));
    
    // Filter tasks by project
    const filtered = mockTasks.filter(task => task.projectid === projectId);
    
    // If no tasks, return empty array with message
    if (filtered.length === 0) {
      console.warn('No tasks found for project:', projectId);
      return [];
    }
    
    return filtered;
  },

  // Get all projects
  getProjects: async (): Promise<Project[]> => {
    await new Promise(resolve => setTimeout(resolve, 300));
    return mockProjects;
  },

  // Get all users
  getUsers: async (): Promise<User[]> => {
    await new Promise(resolve => setTimeout(resolve, 300));
    return mockUsers;
  },

  // Update task (will use real API when available)
  updateTask: async (taskId: number, taskData: Partial<Task>): Promise<Task> => {
    await new Promise(resolve => setTimeout(resolve, 800));
    
    // Find and update mock task
    const taskIndex = mockTasks.findIndex(t => t.id === taskId);
    if (taskIndex === -1) {
      throw new Error('Task not found');
    }
    
    const updatedTask = {
      ...mockTasks[taskIndex],
      ...taskData,
      id: taskId,
    };
    
    mockTasks[taskIndex] = updatedTask;
    console.log('Task updated (mock):', updatedTask);
    
    return updatedTask;
  },

  // Get single task details
  getTask: async (taskId: number): Promise<Task | null> => {
    await new Promise(resolve => setTimeout(resolve, 300));
    const task = mockTasks.find(t => t.id === taskId);
    return task || null;
  },
};