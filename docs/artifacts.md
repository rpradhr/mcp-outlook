# Project Artifacts: Outlook MCP Server

This document lists the key files and components created for the Outlook MCP Server project.

## 1. Core Server & Protocol
- **`server.ts`**: The main entry point. Handles Express middleware, OAuth 2.0 flows, session management, and the MCP SSE (Server-Sent Events) transport layer.
- **`src/mcp/server.ts`**: Contains the MCP server logic using the `@modelcontextprotocol/sdk`. Defines tools (`list_messages`, `send_message`, etc.) and their schemas.

## 2. Integration Services
- **`src/services/outlookService.ts`**: A dedicated service class wrapping the `@microsoft/microsoft-graph-client`. It handles all direct communication with the Microsoft Graph API for mail, calendar, and contacts.

## 3. Frontend (Dashboard)
- **`src/App.tsx`**: A modern React dashboard built with Tailwind CSS and Framer Motion. It provides:
    - OAuth connection status.
    - Login/Logout functionality.
    - Display of the MCP SSE endpoint URL.
    - Visual guide to the server's capabilities.
- **`src/index.css`**: Global styles including custom font imports (Inter, JetBrains Mono) and Tailwind theme configurations.

## 4. Documentation & Design
- **`docs/architecture.md`**: High-level technical design, component diagrams, and tool definitions.
- **`docs/personas.md`**: Detailed user personas (Executive, PM, Consultant, CS Lead) used to drive product requirements.
- **`metadata.json`**: Application metadata including name, description, and required permissions.

## 5. Configuration & Testing
- **`.env.example`**: Template for required environment variables (Microsoft Client ID/Secret, Session Secret).
- **`package.json`**: Project dependencies and scripts (including the `dev` script configured for full-stack mode).
- **`tests/eval-suite.ts`**: A programmatic evaluation suite to verify the Outlook Service tools against a real (or test) access token.

## 6. Infrastructure
- **`vite.config.ts`**: Configured to handle React frontend and proxying.
- **`tsconfig.json`**: TypeScript configuration for the project.
