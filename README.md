# Outlook Procurement Workflow Add-in

A Microsoft Outlook add-in for managing procurement workflows, including RFQ generation, email processing, and quote comparison.

## Features

### RFQ Workflow
- Load open Purchase Requisitions from backend
- Search and select suppliers with match scoring
- Generate RFQs with structured formatting
- Preview and edit RFQ content
- **Send emails directly** (with Graph API) or create drafts

### Email Processing
- Classify incoming emails (quote/clarification/engineer response)
- Extract quote data with structured display
- Handle clarification requests (procurement vs engineering)
- Generate suggested responses
- Forward to engineering team
- Process engineer responses

### Quote Comparison
- View all quotes for an RFQ
- Side-by-side comparison
- Price/delivery analysis

### Folder Management (Graph API)
- Auto-create material-based folder structure
- Organize emails by classification
- Move/copy emails between folders

## Prerequisites

- Node.js 16+ 
- Microsoft 365 account (personal or work)
- Your FastAPI backend running

## Quick Start

### 1. Install Dependencies

```bash
cd outlook-procurement-addin
npm install
```

### 2. Install SSL Certificates

Office Add-ins require HTTPS:

```bash
npx office-addin-dev-certs install
```

### 3. Start Your Backend

```bash
# In your backend directory
uvicorn main:app --reload --port 8000
```

### 4. Start the Add-in Server

```bash
npx http-server . -p 3000 -S -C ~/.office-addin-dev-certs/localhost.crt -K ~/.office-addin-dev-certs/localhost.key --cors
```

### 5. Sideload the Add-in

#### Outlook Web (Recommended for Testing)
1. Go to https://outlook.office.com
2. Open any email
3. Click "..." (More actions) → "Get Add-ins"
4. Click "My add-ins" → "Add a custom add-in" → "Add from file"
5. Select `manifest.xml`

#### Outlook Desktop (Windows)
1. File → Options → Trust Center → Trust Center Settings
2. Trusted Add-in Catalogs
3. Add `https://localhost:3000` as a catalog
4. Restart Outlook
5. Home → Get Add-ins → My Organization

#### Outlook Desktop (Mac)
1. Tools → Get Add-ins
2. My add-ins → Add from file
3. Select `manifest.xml`

## Microsoft Graph API Setup

For full functionality (sending emails, creating folders), you need to configure Graph API:

### Azure App Registration (Already Done)

Your app is registered with:
- **Application (client) ID**: `6492ece4-5f90-4781-bc60-3cb5fc5adb10`
- **Redirect URI**: `https://localhost:3000/taskpane.html`

### Permissions Required

| Permission | Type | Description |
|------------|------|-------------|
| User.Read | Delegated | Sign in and read user profile |
| Mail.ReadWrite | Delegated | Read and write user mail |
| Mail.Send | Delegated | Send mail as user |
| MailboxSettings.ReadWrite | Delegated | Read/write mailbox settings |

### Using Graph API Features

1. Click **"Sign In"** button in the add-in header
2. Microsoft will prompt you to consent to permissions
3. Once signed in, you can:
   - Send emails directly (no draft window)
   - Create and manage folders automatically
   - Move/copy emails between folders

### Without Graph API

If you don't sign in, the add-in still works with limited functionality:
- Creates email drafts (you click Send manually)
- No folder management
- Uses Office.js only

## Project Structure

```
outlook-procurement-addin/
├── manifest.xml              # Add-in configuration
├── package.json              # npm dependencies
├── assets/                   # Icons (16-128px PNG)
├── src/
│   ├── taskpane/
│   │   ├── taskpane.html    # Main UI
│   │   ├── taskpane.css     # Styles
│   │   └── taskpane.js      # Application logic
│   ├── services/
│   │   ├── auth.js          # MSAL authentication
│   │   ├── config.js        # Configuration management
│   │   ├── api-client.js    # Backend API integration
│   │   ├── email-operations.js  # Email send/receive
│   │   └── folder-management.js # Folder operations
│   ├── commands/
│   │   ├── commands.html    # Ribbon commands
│   │   └── commands.js      # Command handlers
│   └── utils/
│       └── helpers.js       # UI utilities
└── test.html                 # Standalone test page
```

## API Endpoints

The add-in integrates with these backend endpoints:

| Method | Endpoint | Description |
|--------|----------|-------------|
| GET | `/api/prs/open` | Get open purchase requisitions |
| POST | `/api/suppliers/search` | Search suppliers for a material |
| POST | `/api/rfqs/generate` | Generate RFQ content |
| POST | `/api/rfqs/finalize` | Finalize and record RFQ |
| POST | `/api/emails/classify` | Classify an email |
| POST | `/api/emails/process` | Process classified email |
| POST | `/api/emails/suggest-response` | Get suggested response |
| POST | `/api/emails/forward-to-engineering` | Forward to engineering |
| POST | `/api/emails/process-engineer-response` | Process engineer reply |
| POST | `/api/emails/extract-quote` | Extract quote data |
| GET | `/api/quotes/{rfq_id}` | Get quotes for comparison |

## Configuration

Click the **Settings** (gear) icon to configure:

- **API Base URL**: Your backend URL (default: `http://localhost:8000`)
- **Engineering Email**: Email for technical queries
- **Auto-create Folders**: Automatically create folder structure

Settings are saved to localStorage.

## Testing Without Outlook

Open `test.html` directly in your browser to:
- Test the UI with mock data
- Verify styling and layout
- Test tab navigation
- No Outlook or backend required

## Troubleshooting

### Add-in doesn't load
- Ensure HTTPS server is running on port 3000
- Check browser console for errors
- Verify SSL certificates are installed

### Sign-in fails
- Check popup blocker settings
- Verify redirect URI matches in Azure: `https://localhost:3000/taskpane.html`
- Check browser console for MSAL errors

### API calls fail
- Verify backend is running on configured port
- Check CORS settings on backend
- Look for errors in browser network tab

### Folders not created
- Must be signed in with Graph API
- Check Mail.ReadWrite permission was granted
- Verify in Azure portal that permissions are consented

## Updating Azure App Registration

If you need to change settings:

1. Go to [Azure Portal](https://portal.azure.com)
2. Navigate to "App registrations"
3. Select "HEXA-OUTLOOK-AGENT"
4. Update settings as needed

### Change Redirect URI
1. Authentication → Platform configurations
2. Edit or add redirect URIs
3. Update `auth.js` if changed

### Add Permissions
1. API permissions → Add a permission
2. Microsoft Graph → Delegated permissions
3. Select permission → Add
4. Grant consent (click "Grant admin consent")

## Production Deployment

For production use:

1. **Update manifest.xml**
   - Change `localhost:3000` to your production URL
   - Update `Id` with a new GUID

2. **Update auth.js**
   - Change `redirectUri` to production URL
   - Update Azure app registration redirect URIs

3. **Backend**
   - Deploy to production server
   - Update API URL in settings

4. **Consider**
   - Code bundling/minification
   - Error logging service
   - Usage analytics

## License

MIT
