/**
 * Outlook MCP Server Evaluation Suite
 * This file contains tests to verify the functionality of the MCP server tools.
 */

import { OutlookService } from '../src/services/outlookService.js';

async function runEval() {
  console.log('🚀 Starting Outlook MCP Evaluation Suite...');
  
  const token = process.env.TEST_ACCESS_TOKEN;
  if (!token) {
    console.error('❌ Error: TEST_ACCESS_TOKEN environment variable is required for evaluation.');
    return;
  }

  const outlook = new OutlookService(token);

  // Test 1: List Messages
  try {
    console.log('Testing list_messages...');
    const messages = await outlook.listMessages(3);
    console.log(`✅ Success: Found ${messages.value.length} messages.`);
  } catch (err: any) {
    console.error(`❌ list_messages failed: ${err.message}`);
  }

  // Test 2: List Events
  try {
    console.log('Testing list_events...');
    const now = new Date();
    const nextWeek = new Date();
    nextWeek.setDate(now.getDate() + 7);
    const events = await outlook.listEvents(now.toISOString(), nextWeek.toISOString());
    console.log(`✅ Success: Found ${events.value.length} events.`);
  } catch (err: any) {
    console.error(`❌ list_events failed: ${err.message}`);
  }

  // Test 3: Search Contacts
  try {
    console.log('Testing search_contacts...');
    const contacts = await outlook.searchContacts('test');
    console.log(`✅ Success: Search returned ${contacts.value.length} contacts.`);
  } catch (err: any) {
    console.error(`❌ search_contacts failed: ${err.message}`);
  }

  console.log('🏁 Evaluation Suite Complete.');
}

if (process.env.RUN_EVAL === 'true') {
  runEval();
}
