# Test Scripts

Direct testing harness for MCP development - bypasses MCP protocol, uses persisted auth.

## Usage

```bash
npm run test:direct
```

## Configuration

Edit `test-harness.ts`:

```typescript
const PRESET = 'tasks';  // Match your --preset
const ORG_MODE = false;  // Match your --org-mode
```

## Testing Your Code

Edit the `runTest()` function:

```typescript
async function runTest(client: GraphClient) {
  // Test any Graph API call
  const result = await client.makeRequest('/me/todo/lists');
  console.log(result);

  // Or import and test your custom tools directly
}
```

## Why Direct Testing?

- MCP servers are long-lived processes
- Code changes don't take effect until restart
- Restarting disconnects Claude and may require re-auth
- Direct testing validates changes WITHOUT restarting
