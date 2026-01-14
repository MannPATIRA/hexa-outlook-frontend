# Quick Deployment Checklist

Follow these steps in order to deploy your add-in.

## 🚀 Quick Start (5 Steps)

### 1. Deploy Frontend to Vercel

```bash
# Install Vercel CLI
npm i -g vercel

# Login
vercel login

# Deploy (from project directory)
cd "outlook-procurement-addin 2"
vercel

# Note your URL (e.g., https://outlook-procurement-addin.vercel.app)
```

### 2. Deploy Backend to Railway

1. Go to [railway.app](https://railway.app) → Sign in with GitHub
2. New Project → Deploy from GitHub repo
3. Select your backend repository
4. Note your URL (e.g., `https://your-app.up.railway.app`)

### 3. Update URLs in Code

```bash
# Update manifest.xml with your Vercel URL
node update-manifest-urls.js https://your-project-name.vercel.app

# Manually update src/services/config.js
# Change: API_BASE_URL to your Railway URL
```

### 4. Update Azure Redirect URI

1. Go to [Azure Portal](https://portal.azure.com)
2. App registrations → HEXA-OUTLOOK-AGENT
3. Authentication → Add redirect URI:
   `https://your-project-name.vercel.app/src/taskpane/taskpane.html`

### 5. Redeploy & Test

```bash
# Redeploy frontend
vercel --prod

# Or push to GitHub if using auto-deploy
```

---

## 📝 Files to Update

After getting your URLs, update these files:

1. **`manifest.xml`** - All `localhost:3000` → Your Vercel URL
   - Use: `node update-manifest-urls.js https://your-url.vercel.app`

2. **`src/services/config.js`** - Line 6:
   ```javascript
   API_BASE_URL: 'https://your-backend.railway.app',
   ```

3. **`src/services/auth.js`** - Already auto-detects (no change needed)

4. **Backend CORS** - Add your Vercel URL to allowed origins

---

## ✅ Verification

- [ ] Frontend deployed to Vercel
- [ ] Backend deployed to Railway/Render
- [ ] manifest.xml URLs updated
- [ ] config.js backend URL updated
- [ ] Azure redirect URI updated
- [ ] Backend CORS updated
- [ ] Redeployed frontend
- [ ] Tested add-in from different device

---

## 🆘 Need Help?

- See `DEPLOYMENT.md` for detailed instructions
- See `BACKEND_DEPLOYMENT.md` for backend-specific help
