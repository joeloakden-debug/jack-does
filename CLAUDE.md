# Jack Does — Project Instructions

## Mandatory: Test Before Push

**Every code change must be visually verified before committing/pushing.**

After making code changes, follow this workflow:

1. **Syntax check** all modified files (`node -c <file>`)
2. **Start the app** using the Preview tool (`preview_start` on `http://localhost:3000`) or verify it's already running
3. **Navigate to the affected page** and take a screenshot to verify the UI renders correctly
4. **Test the specific flow** that was changed — click buttons, fill forms, verify the behavior
5. **Check for errors** — inspect console logs (`preview_console_logs`) and network requests for failures
6. **Only then** commit and push

If the app requires QBO connection or live data that can't be tested locally, at minimum:
- Verify the server starts without errors
- Check that API endpoints return expected shapes (use `preview_eval` or curl)
- Review the payload structure being sent to external APIs (log it)

### Common gotchas to check:
- Tax handling: verify tax lines are detected by both account type AND description
- QBO API payloads: validate structure matches QBO API requirements before sending
- Account type filters: ensure all relevant account types are included in dropdowns
- Month/period logic: verify the correct close period is being used, not current month

## Module Development

When building a new module:
- Write unit tests for core logic (matching, calculations, data transformations)
- Test edge cases: empty data, missing fields, API errors
- Verify the module works with and without QBO connected

## Tech Stack
- **Server**: Node.js (server.js) with Express
- **Frontend**: Vanilla JS (admin/admin.js), HTML (admin/dashboard.html), CSS (admin/admin.css)
- **QBO Integration**: node-quickbooks via qbo-service.js
- **AI**: Claude API (claude-sonnet-4-20250514) for invoice analysis
- **Deployment**: Railway (production at jack-does-production.up.railway.app)
