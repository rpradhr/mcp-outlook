# Outlook MCP Server Architecture

## 1. Overview
The Outlook MCP Server acts as a bridge between AI models (via the Model Context Protocol) and Microsoft Outlook (via Microsoft Graph API). It enables agents to read, search, and compose emails, as well as manage calendar events and contacts.

## 2. Personas & Use Cases
Detailed user personas can be found in [/docs/personas.md](/docs/personas.md). Key personas include:
- **The Executive (Alex)**: High-volume triage and summarization.
- **The Project Manager (Sam)**: Technical tracking and coordination.
- **The Consultant (Elena)**: Client management and billing accuracy.
- **The CS Lead (Jordan)**: Sentiment analysis and response efficiency.

## 3. Technical Stack
- **Runtime**: Node.js (Express)
- **Frontend**: React (for OAuth management and status)
- **API**: Microsoft Graph API
- **Protocol**: Model Context Protocol (MCP)
- **Auth**: OAuth 2.0 (Authorization Code Flow)

## 4. Component Diagram
[AI Client (Claude/Gemini)] <--> [MCP Server (SSE/Stdout)] <--> [MS Graph API] <--> [Outlook Data]

## 5. Tool Definitions
- `list_messages`: List recent emails with optional filters.
- `get_message`: Retrieve full content of a specific email.
- `send_message`: Draft and send a new email.
- `list_events`: List calendar events for a time range.
- `create_event`: Schedule a new meeting.
- `search_contacts`: Find contact details.

## 6. Security
- Tokens are stored in a secure server-side session.
- OAuth scopes are limited to `Mail.ReadWrite`, `Calendars.ReadWrite`, `Contacts.Read`, `User.Read`.
