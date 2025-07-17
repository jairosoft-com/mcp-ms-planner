import http from 'http';
import { IncomingMessage, ServerResponse } from 'http';
import { parse } from 'url';
import { 
  handleGetTasks,
  handleGetTaskDetails,
  handleCreateTask
} from '../services/plannerHandlers.js';

type SseEvent = {
  type: 'task_created' | 'task_updated' | 'task_deleted' | 'error';
  data: any;
  timestamp: string;
};

export class SseServer {
  private server: http.Server;
  private clients: Map<string, ServerResponse> = new Map();
  private port: number;

  constructor(port: number = 3000) {
    this.port = port;
    this.server = http.createServer(this.handleRequest.bind(this));
  }

  private async handleRequest(req: IncomingMessage, res: ServerResponse) {
    const { pathname } = parse(req.url || '/', true);

    // Set CORS headers
    res.setHeader('Access-Control-Allow-Origin', '*');
    res.setHeader('Access-Control-Request-Method', '*');
    res.setHeader('Access-Control-Allow-Methods', 'OPTIONS, GET, POST, PATCH, DELETE');
    res.setHeader('Access-Control-Allow-Headers', '*');

    if (req.method === 'OPTIONS') {
      res.writeHead(200);
      res.end();
      return;
    }

    if (pathname === '/events' && req.method === 'GET') {
      this.handleSseConnection(req, res);
      return;
    }

    // API endpoints
    if (pathname === '/api/planner/tasks') {
      if (req.method === 'GET') {
        await handleGetTasks(req, res);
      } else if (req.method === 'POST') {
        await handleCreateTask(req, res);
      } else {
        res.writeHead(405, { 'Content-Type': 'application/json' });
        res.end(JSON.stringify({ error: 'Method not allowed' }));
      }
      return;
    }

    // Handle task details endpoint: /api/planner/tasks/:id
    if (pathname?.startsWith('/api/planner/tasks/')) {
      if (req.method === 'GET') {
        await handleGetTaskDetails(req, res);
      } else {
        res.writeHead(405, { 'Content-Type': 'application/json' });
        res.end(JSON.stringify({ error: 'Method not allowed' }));
      }
      return;
    }

    res.writeHead(404, { 'Content-Type': 'application/json' });
    res.end(JSON.stringify({ error: 'Endpoint not found' }));
  }

  private handleSseConnection(req: IncomingMessage, res: ServerResponse) {
    // Set headers for SSE
    res.writeHead(200, {
      'Content-Type': 'text/event-stream',
      'Cache-Control': 'no-cache',
      'Connection': 'keep-alive',
    });

    // Generate a unique ID for this client
    const clientId = Date.now().toString();
    
    // Add client to the map
    this.clients.set(clientId, res);

    // Send initial connection message
    this.sendEvent(res, 'connected', { 
      clientId,
      message: 'Connected to planner server',
      timestamp: new Date().toISOString()
    });

    // Handle client disconnect
    req.on('close', () => {
      this.clients.delete(clientId);
      console.log(`Client ${clientId} disconnected`);
    });
  }

  // Broadcast a task event to all connected clients
  public broadcastTaskEvent(eventType: 'task_created' | 'task_updated' | 'task_deleted', task: any) {
    this.broadcast(eventType, {
      ...task,
      eventType,
      timestamp: new Date().toISOString()
    });
  }

  public sendEvent(res: ServerResponse, event: string, data: any) {
    try {
      const message = `event: ${event}\ndata: ${JSON.stringify(data)}\n\n`;
      res.write(message);
      // Ensure the message is sent immediately
      if ((res as any).flush) {
        (res as any).flush();
      }
    } catch (error) {
      console.error('Error sending SSE event:', error);
    }
  }

  public broadcast(eventType: string, data: any) {
    const event: SseEvent = {
      type: eventType as any,
      data,
      timestamp: new Date().toISOString()
    };

    // Send to all connected clients
    for (const [clientId, client] of this.clients.entries()) {
      try {
        this.sendEvent(client, eventType, event);
      } catch (error) {
        console.error(`Error sending to client ${clientId}:`, error);
        this.clients.delete(clientId);
      }
    }
  }

  public async start(): Promise<void> {
    return new Promise<void>((resolve, reject) => {
      this.server.on('error', (error) => {
        console.error('Server error:', error);
        reject(error);
      });

      this.server.listen(this.port, () => {
        console.log(`Planner SSE Server running on http://localhost:${this.port}`);
        console.log(`SSE endpoint: http://localhost:${this.port}/events`);
        console.log(`Planner API endpoint: http://localhost:${this.port}/api/planner/tasks`);
        resolve();
      });
    });
  }

  public stop(): Promise<void> {
    return new Promise((resolve, reject) => {
      // Close all client connections
      for (const [clientId, client] of this.clients.entries()) {
        try {
          client.end();
        } catch (error) {
          console.error(`Error closing client ${clientId}:`, error);
        }
      }
      this.clients.clear();

      // Close the server
      this.server.close((error) => {
        if (error) {
          console.error('Error stopping server:', error);
          reject(error);
        } else {
          console.log('Planner SSE Server stopped');
          resolve();
        }
      });
    });
  }
}

export default SseServer;
