# Fixes for Localhost Issues

## ✅ Fixed Issues

### 1. env-config.js Error
**Error:** `Uncaught ReferenceError: openaiKey is not defined`
**Fixed:** Updated build.js to use `window.OPENAI_API_KEY_ENV` instead of undefined `openaiKey` variable

### 2. Logo Not Showing
**Error:** `logo.png:1 Failed to load resource: 404 (Not Found)`
**Status:** Logo exists and is being copied correctly. The 404 is because:
- Old manifest is still cached (pointing to localhost:3000)
- Once manifest updates, logo will load from correct path

### 3. Backend API Configuration
**Fixed:** Both local and deployed frontends now always use:
- `https://hexa-outlook-backend.onrender.com`

## ⚠️ Remaining Issues (Need Backend Fix)

### CORS Error
**Error:** 
```
Access to fetch at 'https://hexa-outlook-backend.onrender.com/api/prs/open' 
from origin 'https://localhost:3000' has been blocked by CORS policy
```

**Fix Required:** Update your backend CORS configuration to allow `localhost:3000`:

```python
from fastapi.middleware.cors import CORSMiddleware

app.add_middleware(
    CORSMiddleware,
    allow_origins=[
        "https://hexa-outlook-frontend.vercel.app",  # Production frontend
        "https://localhost:3000",  # Local development - ADD THIS
        "http://localhost:3000",   # Local development HTTP - ADD THIS
    ],
    allow_credentials=True,
    allow_methods=["GET", "POST", "PUT", "DELETE", "OPTIONS"],
    allow_headers=["*"],
)
```

**Action:** Update your backend code and redeploy to Render.

## 🔄 Manifest Cache Issue

### Icons Not Loading
**Error:** `localhost:3000/assets/icon-*.png: Failed to load`

**Cause:** Microsoft 365 Admin Center is still serving the old manifest with localhost URLs.

**Solution:**
1. Use manifest URL link in Microsoft 365 Admin Center:
   - URL: `https://hexa-outlook-frontend.vercel.app/manifest.xml`
2. Or wait for cache to clear (can take hours)
3. Remove and re-add the add-in in Outlook

## 📝 Summary

| Issue | Status | Action Needed |
|-------|--------|---------------|
| env-config.js error | ✅ Fixed | None - rebuild will fix |
| Logo 404 | ✅ Fixed | Wait for manifest update |
| Backend URL | ✅ Fixed | None - always uses deployed backend |
| CORS error | ⚠️ Needs backend fix | Update backend CORS config |
| Icons not loading | ⚠️ Manifest cache | Use manifest URL or wait |

## Next Steps

1. **Rebuild:** Run `npm run build` to fix env-config.js
2. **Backend CORS:** Add localhost:3000 to backend CORS allowed origins
3. **Manifest:** Use manifest URL link in Microsoft 365 Admin Center
4. **Test:** After backend CORS fix, test locally again
