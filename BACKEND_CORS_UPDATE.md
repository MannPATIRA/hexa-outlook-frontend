# Backend CORS Configuration Update

## Your Backend URL
**Backend:** `https://hexa-outlook-backend.onrender.com`

## Your Frontend URL
**Frontend:** `https://hexa-outlook-frontend.vercel.app` (or your actual Vercel URL)

## Required Backend Changes

You need to update your backend's CORS configuration to allow requests from your Vercel frontend.

### Update CORS in your FastAPI backend:

```python
from fastapi.middleware.cors import CORSMiddleware

app.add_middleware(
    CORSMiddleware,
    allow_origins=[
        "https://hexa-outlook-frontend.vercel.app",  # Your Vercel frontend
        "https://hexa-outlook-frontend-git-main-manns-projects-a4a9cb77.vercel.app",  # Preview deployments
        "https://localhost:3000",  # For local development
        "http://localhost:3000",   # For local development
    ],
    allow_credentials=True,
    allow_methods=["GET", "POST", "PUT", "DELETE", "OPTIONS"],
    allow_headers=["*"],
)
```

### Important Notes:

1. **Replace the Vercel URL** with your actual deployment URL if it's different
2. **Add all Vercel preview URLs** - Vercel creates different URLs for each branch/deployment
3. **Keep localhost** for local development
4. **Deploy the backend** after making these changes

### How to Find Your Exact Vercel URL:

1. Go to your Vercel dashboard
2. Select your project
3. Check the "Domains" section in deployment details
4. You'll see URLs like:
   - `hexa-outlook-frontend.vercel.app` (production)
   - `hexa-outlook-frontend-*.vercel.app` (preview deployments)

Add all of these to the `allow_origins` list.

### Testing:

After updating CORS, test by:
1. Opening your deployed frontend
2. Opening browser DevTools → Network tab
3. Try to make an API call (e.g., load PRs)
4. Check if CORS errors appear in console

If you see CORS errors, the backend URL might not match exactly - double-check the URL in the error message.
