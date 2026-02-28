import { Server } from '@modelcontextprotocol/sdk/server/index.js';
import { SSEServerTransport } from '@modelcontextprotocol/sdk/server/sse.js';
import { 
  CallToolRequestSchema, 
  ListToolsRequestSchema,
  ErrorCode,
  McpError
} from '@modelcontextprotocol/sdk/types.js';
import { OutlookService } from '../services/outlookService.js';

export class OutlookMcpServer {
  private server: Server;

  constructor() {
    this.server = new Server(
      {
        name: 'outlook-mcp-server',
        version: '1.0.0',
      },
      {
        capabilities: {
          tools: {},
        },
      }
    );

    this.setupHandlers();
  }

  private setupHandlers() {
    this.server.setRequestHandler(ListToolsRequestSchema, async () => ({
      tools: [
        {
          name: 'list_messages',
          description: 'List recent emails from the inbox',
          inputSchema: {
            type: 'object',
            properties: {
              top: { type: 'number', description: 'Number of messages to retrieve (default 10)' },
            },
          },
        },
        {
          name: 'get_message',
          description: 'Get the full content of a specific email',
          inputSchema: {
            type: 'object',
            properties: {
              id: { type: 'string', description: 'The ID of the message' },
            },
            required: ['id'],
          },
        },
        {
          name: 'send_message',
          description: 'Send a new email',
          inputSchema: {
            type: 'object',
            properties: {
              to: { type: 'string', description: 'Recipient email address' },
              subject: { type: 'string', description: 'Email subject' },
              body: { type: 'string', description: 'Email body content' },
            },
            required: ['to', 'subject', 'body'],
          },
        },
        {
          name: 'list_events',
          description: 'List calendar events for a given time range',
          inputSchema: {
            type: 'object',
            properties: {
              start: { type: 'string', description: 'Start time (ISO 8601)' },
              end: { type: 'string', description: 'End time (ISO 8601)' },
            },
            required: ['start', 'end'],
          },
        },
        {
          name: 'create_event',
          description: 'Create a new calendar event',
          inputSchema: {
            type: 'object',
            properties: {
              subject: { type: 'string', description: 'Event subject' },
              start: { type: 'string', description: 'Start time (ISO 8601)' },
              end: { type: 'string', description: 'End time (ISO 8601)' },
              location: { type: 'string', description: 'Event location' },
            },
            required: ['subject', 'start', 'end'],
          },
        },
      ],
    }));

    this.server.setRequestHandler(CallToolRequestSchema, async (request) => {
      // Note: In a real implementation, we'd need to pass the access token here.
      // Since MCP is stateless per request but our server has session, 
      // we'll handle the token retrieval in the transport layer or via a factory.
      throw new McpError(ErrorCode.InternalError, 'Token required for tool execution');
    });
  }

  public async connect(transport: SSEServerTransport) {
    await this.server.connect(transport);
  }

  public getMcpServer() {
    return this.server;
  }
}
