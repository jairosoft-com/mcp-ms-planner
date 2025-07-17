import { IncomingMessage, ServerResponse } from 'http';
import { URLSearchParams } from 'url';
import { PlannerTask } from '../interfaces/plannerTask.js';
import { fetchPlannerTasks, getTaskDetails, createPlannerTask } from './plannerService.js';

export async function handleGetTasks(req: IncomingMessage, res: ServerResponse) {
  try {
    const url = new URL(req.url || '/', `http://${req.headers.host}`);
    const status = url.searchParams.get('status') as 'notStarted' | 'inProgress' | 'completed' | undefined;
    
    // Get the access token from the Authorization header
    const authHeader = req.headers.authorization || '';
    const accessToken = authHeader.startsWith('Bearer ') ? authHeader.slice(7) : '';
    
    if (!accessToken) {
      res.writeHead(401, { 'Content-Type': 'application/json' });
      res.end(JSON.stringify({ error: 'Unauthorized - No access token provided' }));
      return;
    }

    const { tasks, count, hasMore } = await fetchPlannerTasks({ status, accessToken });
    
    res.writeHead(200, { 'Content-Type': 'application/json' });
    res.end(JSON.stringify({ tasks, count, hasMore }));
  } catch (error: any) {
    console.error('Error fetching tasks:', error);
    res.writeHead(500, { 'Content-Type': 'application/json' });
    res.end(JSON.stringify({ 
      error: 'Failed to fetch tasks', 
      details: error.message 
    }));
  }
}

export async function handleGetTaskDetails(req: IncomingMessage, res: ServerResponse) {
  try {
    const url = new URL(req.url || '/', `http://${req.headers.host}`);
    const taskId = url.pathname.split('/').pop();
    
    if (!taskId) {
      res.writeHead(400, { 'Content-Type': 'application/json' });
      res.end(JSON.stringify({ error: 'Task ID is required' }));
      return;
    }
    
    // Get the access token from the Authorization header
    const authHeader = req.headers.authorization || '';
    const accessToken = authHeader.startsWith('Bearer ') ? authHeader.slice(7) : '';
    
    if (!accessToken) {
      res.writeHead(401, { 'Content-Type': 'application/json' });
      res.end(JSON.stringify({ error: 'Unauthorized - No access token provided' }));
      return;
    }

    const task = await getTaskDetails(taskId, accessToken);
    
    if (!task) {
      res.writeHead(404, { 'Content-Type': 'application/json' });
      res.end(JSON.stringify({ error: 'Task not found' }));
      return;
    }
    
    res.writeHead(200, { 'Content-Type': 'application/json' });
    res.end(JSON.stringify(task));
  } catch (error: any) {
    console.error('Error fetching task details:', error);
    res.writeHead(500, { 'Content-Type': 'application/json' });
    res.end(JSON.stringify({ 
      error: 'Failed to fetch task details', 
      details: error.message 
    }));
  }
}

export async function handleCreateTask(req: IncomingMessage, res: ServerResponse) {
  try {
    // Get the access token from the Authorization header
    const authHeader = req.headers.authorization || '';
    const accessToken = authHeader.startsWith('Bearer ') ? authHeader.slice(7) : '';
    
    if (!accessToken) {
      res.writeHead(401, { 'Content-Type': 'application/json' });
      res.end(JSON.stringify({ error: 'Unauthorized - No access token provided' }));
      return;
    }
    
    // Read and parse the request body
    let body = '';
    for await (const chunk of req) {
      body += chunk.toString();
    }
    
    const taskData = JSON.parse(body || '{}');
    
    // Validate required fields
    if (!taskData.title || !taskData.planId || !taskData.bucketId) {
      res.writeHead(400, { 'Content-Type': 'application/json' });
      res.end(JSON.stringify({ 
        error: 'Missing required fields', 
        required: ['title', 'planId', 'bucketId'] 
      }));
      return;
    }
    
    const newTask = await createPlannerTask({
      ...taskData,
      accessToken
    });
    
    res.writeHead(201, { 'Content-Type': 'application/json' });
    res.end(JSON.stringify(newTask));
  } catch (error: any) {
    console.error('Error creating task:', error);
    res.writeHead(500, { 'Content-Type': 'application/json' });
    res.end(JSON.stringify({ 
      error: 'Failed to create task', 
      details: error.message 
    }));
  }
}

export default {
  handleGetTasks,
  handleGetTaskDetails,
  handleCreateTask
};
