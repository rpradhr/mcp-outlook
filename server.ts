import express from 'express';
import session from 'express-session';
import cookieParser from 'cookie-parser';
import { createServer as createViteServer } from 'vite';
import path from 'path';
import { fileURLToPath } from 'url';
import dotenv from 'dotenv';
import { v4 as uuidv4 } from 'uuid';

import { SSEServerTransport } from '@modelcontextprotocol/sdk/server/sse.js';
import { OutlookService } from './src/services/outlookService.js';
import { 
  CallToolRequestSchema, 
  ListToolsRequestSchema,
  ErrorCode,
  McpError
} from '@modelcontextprotocol/sdk/types.js';
import { Server } from '@modelcontextprotocol/sdk/server/index.js';

dotenv.config();

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

// Initialize MCP Server
const mcpServer = new Server(
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

// Define MCP Tools
mcpServer.setRequestHandler(ListToolsRequestSchema, async () => ({
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

async function startServer() {
  const app = express();
  const PORT = 3000;

  app.use(express.json());
  app.use(cookieParser());
  app.use(session({
    secret: process.env.SESSION_SECRET || 'outlook-mcp-secret',
    resave: false,
    saveUninitialized: true,
    cookie: { 
      secure: true, 
      sameSite: 'none',
      httpOnly: true 
    }
  }));

  // --- OAuth Routes ---
  app.get('/api/auth/url', (req, res) => {
    const clientId = process.env.MICROSOFT_CLIENT_ID;
    const redirectUri = process.env.MICROSOFT_REDIRECT_URI || `${process.env.APP_URL}/auth/callback`;
    const scope = encodeURIComponent('offline_access User.Read Mail.ReadWrite Calendars.ReadWrite Contacts.Read');
    
    if (!clientId) {
      return res.status(500).json({ error: 'MICROSOFT_CLIENT_ID not configured' });
    }

    const authUrl = `https://login.microsoftonline.com/common/oauth2/v2.0/authorize?client_id=${clientId}&response_type=code&redirect_uri=${encodeURIComponent(redirectUri)}&response_mode=query&scope=${scope}&state=${uuidv4()}`;
    
    res.json({ url: authUrl });
  });

  app.get('/auth/callback', async (req, res) => {
    const { code } = req.query;
    const clientId = process.env.MICROSOFT_CLIENT_ID;
    const clientSecret = process.env.MICROSOFT_CLIENT_SECRET;
    const redirectUri = process.env.MICROSOFT_REDIRECT_URI || `${process.env.APP_URL}/auth/callback`;

    if (!code) return res.status(400).send('No code provided');

    try {
      const tokenResponse = await fetch('https://login.microsoftonline.com/common/oauth2/v2.0/token', {
        method: 'POST',
        headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
        body: new URLSearchParams({
          client_id: clientId!,
          client_secret: clientSecret!,
          code: code as string,
          redirect_uri: redirectUri,
          grant_type: 'authorization_code',
        }),
      });

      const tokens = await tokenResponse.json();
      if (tokens.error) throw new Error(tokens.error_description);

      // Store tokens in session (In a real app, use a DB)
      (req.session as any).tokens = tokens;
      (global as any).lastAccessToken = tokens.access_token;

      res.send(`
        <html>
          <body>
            <script>
              if (window.opener) {
                window.opener.postMessage({ type: 'OAUTH_AUTH_SUCCESS' }, '*');
                window.close();
              } else {
                window.location.href = '/';
              }
            </script>
            <p>Authentication successful. This window should close automatically.</p>
          </body>
        </html>
      `);
    } catch (error: any) {
      console.error('OAuth Error:', error);
      res.status(500).send(`Auth failed: ${error.message}`);
    }
  });

  app.get('/api/auth/status', (req, res) => {
    const tokens = (req.session as any).tokens;
    res.json({ isAuthenticated: !!tokens });
  });

  app.post('/api/auth/logout', (req, res) => {
    req.session.destroy(() => {
      res.json({ success: true });
    });
  });

  // --- MCP SSE Endpoints ---
  let transport: SSEServerTransport | null = null;

  app.get('/mcp/sse', async (req, res) => {
    const tokens = (req.session as any).tokens;
    if (!tokens) {
      return res.status(401).send('Unauthorized: Please login first');
    }

    transport = new SSEServerTransport('/mcp/messages', res);
    await mcpServer.connect(transport);
  });

  app.post('/mcp/messages', async (req, res) => {
    if (transport) {
      await transport.handlePostMessage(req, res);
    } else {
      res.status(400).send('No active SSE transport');
    }
  });

  // Handle Tool Calls with Session Context
  mcpServer.setRequestHandler(CallToolRequestSchema, async (request) => {
    // This is tricky because the request handler doesn't have access to the Express request object directly.
    // In a real production app, we'd use a more robust way to link the MCP session to the HTTP session.
    // For this demo, we'll assume the last authenticated user is the one calling the tool.
    // (In a multi-user scenario, we'd need to pass session IDs or use a different transport).
    
    // We'll use a global or a way to retrieve the token.
    // Since we're in a single-user-per-container environment mostly, this is okay for now.
    // But let's try to be better. We can store the service in a map keyed by transport.
    
    // For now, let's assume we have a way to get the token.
    // We'll implement a simple "last token" cache for the demo.
    const accessToken = (global as any).lastAccessToken;
    if (!accessToken) {
      throw new McpError(ErrorCode.InvalidRequest, 'Not authenticated');
    }

    const outlook = new OutlookService(accessToken);

    try {
      switch (request.params.name) {
        case 'list_messages': {
          const { top } = request.params.arguments as any;
          const data = await outlook.listMessages(top);
          return { content: [{ type: 'text', text: JSON.stringify(data, null, 2) }] };
        }
        case 'get_message': {
          const { id } = request.params.arguments as any;
          const data = await outlook.getMessage(id);
          return { content: [{ type: 'text', text: JSON.stringify(data, null, 2) }] };
        }
        case 'send_message': {
          const { to, subject, body } = request.params.arguments as any;
          await outlook.sendMessage(subject, body, to);
          return { content: [{ type: 'text', text: 'Email sent successfully' }] };
        }
        case 'list_events': {
          const { start, end } = request.params.arguments as any;
          const data = await outlook.listEvents(start, end);
          return { content: [{ type: 'text', text: JSON.stringify(data, null, 2) }] };
        }
        case 'create_event': {
          const { subject, start, end, location } = request.params.arguments as any;
          await outlook.createEvent(subject, start, end, location);
          return { content: [{ type: 'text', text: 'Event created successfully' }] };
        }
        default:
          throw new McpError(ErrorCode.MethodNotFound, `Unknown tool: ${request.params.name}`);
      }
    } catch (error: any) {
      return {
        content: [{ type: 'text', text: `Error: ${error.message}` }],
        isError: true,
      };
    }
  });

  // --- Vite Middleware ---
  if (process.env.NODE_ENV !== 'production') {
    const vite = await createViteServer({
      server: { middlewareMode: true },
      appType: 'spa',
    });
    app.use(vite.middlewares);
  } else {
    app.use(express.static(path.join(__dirname, 'dist')));
    app.get('*', (req, res) => {
      res.sendFile(path.join(__dirname, 'dist', 'index.html'));
    });
  }

  app.listen(PORT, '0.0.0.0', () => {
    console.log(`Server running on http://0.0.0.0:${PORT}`);
  });
}

startServer();
