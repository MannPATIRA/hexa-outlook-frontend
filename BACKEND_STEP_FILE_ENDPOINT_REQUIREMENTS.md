# Backend STEP File Endpoint Requirements

## 🚨 CRITICAL ISSUE
**Status**: STEP files are returning **HTTP 404 Not Found** from backend  
**Impact**: STEP files cannot be attached to email drafts  
**Priority**: HIGH - Blocks core functionality

## Quick Fix Required
The backend must implement/fix the file endpoint:
```
GET /api/files/{filename}
```

## Issue Summary
The frontend is receiving **HTTP 404 Not Found** errors when trying to fetch STEP files from the backend. The files are listed in the RFQ API response but are not accessible via the file endpoint.

## Current Error
```
Failed to fetch file model_PR001.step: HTTP 404 - {"detail":"Not Found"}
Failed to fetch file assembly_PR001.step: HTTP 404 - {"detail":"Not Found"}
```

## Frontend Request Details

### Base URL
```
https://hexa-outlook-backend.onrender.com
```

### API Prefix
```
/api
```

### Endpoint Format
The frontend makes requests to **two possible endpoint formats**:

1. **Without RFQ ID** (most common):
   ```
   GET /api/files/{filename}
   ```
   Example: `GET /api/files/model_PR001.step`

2. **With RFQ ID** (when provided):
   ```
   GET /api/files/rfq/{rfqId}/{filename}
   ```
   Example: `GET /api/files/rfq/RFQ123/model_PR001.step`

### Request Headers
```
Accept: application/octet-stream
```

### Full URL Examples
- `https://hexa-outlook-backend.onrender.com/api/files/model_PR001.step`
- `https://hexa-outlook-backend.onrender.com/api/files/assembly_PR001.step`

## Expected Response

### Success Response (HTTP 200)
- **Content-Type**: `application/octet-stream` (or appropriate binary MIME type)
- **Body**: Binary file content (the STEP file bytes)
- **Headers**: 
  - `Content-Length`: Size of file in bytes
  - `Content-Type`: `application/octet-stream` or `application/octet-stream; charset=binary`

### Error Response (Current - HTTP 404)
```json
{"detail":"Not Found"}
```

## What the Backend Needs to Fix

### 1. Verify File Storage Location
- Confirm STEP files are actually stored on the Render server
- Check if files are in a `files/` directory or different location
- Verify file naming matches what's in the RFQ API response

### 2. Implement/Verify Endpoint
The backend must implement a GET endpoint at:
```
GET /api/files/{filename}
```

That:
- Accepts the filename (URL-encoded)
- Returns the file as binary data with `application/octet-stream` content type
- Returns HTTP 404 if file doesn't exist
- Handles URL encoding/decoding of filenames correctly

### 3. RFQ-Specific Endpoint (Currently Used)
The frontend **does pass RFQ IDs** when fetching attachments for specific RFQs. The backend should implement:
```
GET /api/files/rfq/{rfqId}/{filename}
```

**Note**: The frontend will try this endpoint when `rfqId` is provided. If this endpoint doesn't exist or returns 404, the frontend will fall back to the general `/api/files/{filename}` endpoint.

**Example**: When fetching files for RFQ with ID `RFQ123`, the frontend will request:
- `GET /api/files/rfq/RFQ123/model_PR001.step`

### 4. CORS Configuration
Ensure the file endpoint allows CORS requests from:
- Outlook web app domains
- The frontend deployment domain

### 5. File Path Resolution
The backend should:
- Decode the URL-encoded filename
- Look up the file in the storage location
- Return the file if found, or 404 if not found

## Testing

### Test URLs (for backend team to verify)
1. `https://hexa-outlook-backend.onrender.com/api/files/model_PR001.step`
2. `https://hexa-outlook-backend.onrender.com/api/files/assembly_PR001.step`

### Expected Test Results
- **Success**: File downloads as binary data
- **Failure**: HTTP 404 with JSON error (current behavior)

## Files Affected

### Frontend Code Locations
- `src/services/api-client.js` - File fetching logic (lines 247-334)
- `src/utils/attachments.js` - Attachment preparation (lines 348-370)
- `src/services/config.js` - API URL configuration

### Backend Endpoints Needed
- `GET /api/files/{filename}` - **REQUIRED**
- `GET /api/files/rfq/{rfqId}/{filename}` - **OPTIONAL** (if RFQ-specific paths are used)

## Additional Notes

1. **File Types**: The frontend handles STEP files (`.step`, `.stp`) as binary files with `application/octet-stream` content type.

2. **Error Handling**: The frontend now shows detailed error messages including HTTP status codes, so any backend errors will be visible to users.

3. **File Size**: No specific size limits are enforced in the frontend, but ensure backend can handle typical STEP file sizes (often several MB).

4. **Authentication**: Verify if the file endpoint requires authentication. If so, ensure the frontend's authentication tokens are being sent correctly.

## Next Steps for Backend Team

1. ✅ Check if STEP files are deployed to Render server
2. ✅ Verify file storage location and naming
3. ✅ Implement/fix the `/api/files/{filename}` endpoint
4. ✅ Test endpoint with actual STEP filenames from RFQ responses
5. ✅ Verify CORS configuration allows file downloads
6. ✅ Confirm files are accessible via the exact URLs the frontend is requesting

## Contact
If the backend team needs clarification on the frontend requirements, please refer to this document or check the frontend code in the files listed above.
