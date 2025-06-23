// Type definitions for @modelcontextprotocol/sdk
declare module '@modelcontextprotocol/sdk/server/mcp' {
  export interface RequestHandler {
    send(response: any): void;
  }

  export interface Request {
    jsonrpc: string;
    method: string;
    params?: Record<string, any>;
    id?: string | number | null;
  }

  export interface Response {
    jsonrpc: string;
    result?: any;
    error?: {
      code: number;
      message: string;
      data?: any;
    };
    id: string | number | null;
  }

  export class McpServer {
    constructor(options: { name: string; version: string });
    
    tool(
      name: string,
      description: string,
      paramsSchema: any,
      handler: (args: any) => Promise<{ content: Array<{ type: string; text: string }> }>
    ): void;
    
    connect(transport: any): Promise<void>;
    disconnect(): Promise<void>;
    handleRequest(request: Request, handler: RequestHandler): Promise<void>;
  }
}

declare module '@modelcontextprotocol/sdk/server/stdio' {
  export class StdioServerTransport {
    constructor();
  }
}
