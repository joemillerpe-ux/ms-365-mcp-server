/**
 * Direct test harness - bypasses MCP, uses persisted auth tokens.
 *
 * Usage:
 *   npx tsx test-scripts/test-harness.ts
 *   npm run test:direct
 *
 * Edit the `runTest()` function to test your code changes.
 */

import AuthManager, { buildScopesFromEndpoints } from '../src/auth.js';
import GraphClient from '../src/graph-client.js';
import { getSecrets } from '../src/secrets.js';
import { getCombinedPresetPattern } from '../src/tool-categories.js';

// ============================================================
// CONFIGURATION - Match your MCP server settings
// ============================================================
const PRESET = 'tasks';  // Same as --preset tasks
const ORG_MODE = false;  // Set true if using --org-mode
// ============================================================

async function setup() {
  console.log(`Loading with preset: ${PRESET}`);

  // Build the same tool filter pattern as the MCP server
  const enabledToolsPattern = getCombinedPresetPattern([PRESET]);
  console.log(`  Tool filter: ${enabledToolsPattern}`);

  // Build scopes matching MCP server configuration
  const scopes = buildScopesFromEndpoints(ORG_MODE, enabledToolsPattern);
  console.log(`  Scopes: ${scopes.join(', ')}`);

  console.log('Creating AuthManager...');
  const auth = await AuthManager.create(scopes);

  console.log('Loading token cache...');
  await auth.loadTokenCache();

  // Check what accounts are available
  const accounts = await auth.listAccounts();
  console.log('  Cached accounts:', accounts.length);
  accounts.forEach((a, i) => console.log(`    [${i}] ${a.username}`));

  if (accounts.length === 0) {
    console.error('\nNo cached accounts! Run the MCP server first to authenticate.');
    console.log('  npx tsx src/index.ts --preset tasks');
    console.log('  Then use verify-login or login tool');
    process.exit(1);
  }

  const secrets = await getSecrets();
  const client = new GraphClient(auth, secrets);

  console.log('Setup complete.\n');
  return { auth, client, secrets };
}

// ============================================================
// EDIT THIS FUNCTION TO TEST YOUR CODE
// ============================================================
async function runTest(client: GraphClient) {
  // Example: List To-Do lists
  const result = await client.makeRequest('/me/todo/lists');
  console.log(JSON.stringify(result, null, 2));

  // Example: Test your conversion tool
  // import { registerConversionTools } from '../src/conversion-tools.js';
  // ...
}

// ============================================================

async function main() {
  try {
    const { client } = await setup();
    await runTest(client);
  } catch (error) {
    console.error('Error:', error);
    process.exit(1);
  }
}

main();
