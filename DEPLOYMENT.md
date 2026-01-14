# Deployment Guide

This guide will help you deploy your Outlook add-in so it works from any device without running a local server.

## Overview

- **Frontend (Add-in)**: Deploy to **Vercel** (free, easy, automatic HTTPS)
- **Backend (FastAPI)**: Deploy to **Railway** or **Render** (both have free tiers)

---

## Part 1: Deploy Frontend to Vercel

### Step 1: Prepare Your Code

1. **Update `manifest.xml`** - Replace all `localhost:3000` URLs with your Vercel URL
   - After deployment, your Vercel URL will be: `https://your-project-name.vercel.app`
   - We'll update this after deployment

2. **Update `src/services/auth.js`** - Update the redirect URI
   - Change `redirectUri: 'https://localhost:3000/src/taskpane/taskpane.html'`
   - To: `redirectUri: 'https://your-project-name.vercel.app/src/taskpane/taskpane.html'`

3. **Update `src/services/config.js`** - Set production backend URL
   - Change `API_BASE_URL: 'http://localhost:8000'`
   - To: `API_BASE_URL: 'https://your-backend-url.railway.app'` (or your Render URL)

### Step 2: Deploy to Vercel

#### Option A: Using Vercel CLI (Recommended)

1. **Install Vercel CLI**:
   ```bash
   npm i -g vercel
   ```

2. **Login to Vercel**:
   ```bash
   vercel login
   ```

3. **Deploy**:
   ```bash
   cd "outlook-procurement-addin 2"
   vercel
   ```

4. **Follow the prompts**:
   - Set up and deploy? **Yes**
   - Which scope? (select your account)
   - Link to existing project? **No**
   - Project name? (press Enter for default or enter a name)
   - Directory? (press Enter for current directory)
   - Override settings? **No**

5. **Note your deployment URL** (e.g., `https://outlook-procurement-addin.vercel.app`)

6. **Update URLs** in your code with the actual Vercel URL, then redeploy:
   ```bash
   vercel --prod
   ```

#### Option B: Using GitHub (Easier for Updates)

1. **Push your code to GitHub**:
   ```bash
   git init
   git add .
   git commit -m "Initial commit"
   git remote add origin https://github.com/yourusername/outlook-procurement-addin.git
   git push -u origin main
   ```

2. **Go to [vercel.com](https://vercel.com)** and sign in with GitHub

3. **Click "Add New Project"**

4. **Import your repository**

5. **Configure**:
   - Framework Preset: **Other**
   - Root Directory: `./`
   - Build Command: (leave empty)
   - Output Directory: `./`

6. **Deploy**

7. **After deployment**, update your URLs and push again - Vercel will auto-deploy

### Step 3: Update Azure App Registration

1. Go to [Azure Portal](https://portal.azure.com)
2. Navigate to **App registrations** → **HEXA-OUTLOOK-AGENT**
3. Go to **Authentication**
4. Under **Platform configurations**, add a new **Single-page application**:
   - Redirect URI: `https://your-project-name.vercel.app/src/taskpane/taskpane.html`
5. Click **Save**

---

## Part 2: Deploy Backend to Railway

### Step 1: Prepare Your Backend

Your backend should be a separate FastAPI application. Make sure it has:

1. **CORS enabled** for your Vercel domain:
   ```python
   from fastapi.middleware.cors import CORSMiddleware
   
   app.add_middleware(
       CORSMiddleware,
       allow_origins=[
           "https://your-project-name.vercel.app",
           "https://localhost:3000"  # For local testing
       ],
       allow_credentials=True,
       allow_methods=["*"],
       allow_headers=["*"],
   )
   ```

2. **A `requirements.txt`** file with all dependencies

3. **A startup command** in `Procfile` or `railway.json`:
   ```
   web: uvicorn main:app --host 0.0.0.0 --port $PORT
   ```

### Step 2: Deploy to Railway

1. **Go to [railway.app](https://railway.app)** and sign in with GitHub

2. **Click "New Project"** → **"Deploy from GitHub repo"**

3. **Select your backend repository**

4. **Railway will auto-detect** your Python app and deploy it

5. **Get your backend URL**:
   - Go to your project → Settings → Domains
   - Railway provides a default domain like: `your-app.up.railway.app`
   - Or add a custom domain

6. **Set environment variables** (if needed):
   - Go to Variables tab
   - Add any API keys, database URLs, etc.

### Step 3: Update Frontend Config

Update `src/services/config.js` in your frontend:
```javascript
API_BASE_URL: 'https://your-app.up.railway.app',
```

Then redeploy to Vercel.

---

## Part 3: Deploy Backend to Render (Alternative)

If you prefer Render over Railway:

1. **Go to [render.com](https://render.com)** and sign in

2. **Click "New +"** → **"Web Service"**

3. **Connect your GitHub repository** (backend)

4. **Configure**:
   - Name: `procurement-backend`
   - Environment: **Python 3**
   - Build Command: `pip install -r requirements.txt`
   - Start Command: `uvicorn main:app --host 0.0.0.0 --port $PORT`

5. **Click "Create Web Service"**

6. **Get your URL**: `https://your-app.onrender.com`

7. **Update frontend config** and redeploy

---

## Part 4: Update Manifest and Redeploy

After you have both URLs:

1. **Update `manifest.xml`** - Replace ALL instances of `localhost:3000`:
   ```bash
   # Find and replace in your editor
   https://localhost:3000 → https://your-project-name.vercel.app
   ```

2. **Update `src/services/auth.js`**:
   ```javascript
   redirectUri: 'https://your-project-name.vercel.app/src/taskpane/taskpane.html',
   ```

3. **Update `src/services/config.js`**:
   ```javascript
   API_BASE_URL: 'https://your-backend-url.railway.app',  // or .onrender.com
   ```

4. **Redeploy to Vercel**:
   ```bash
   vercel --prod
   ```
   Or push to GitHub if using auto-deploy

---

## Part 5: Sideload the Add-in

### For Outlook Web:

1. Go to [outlook.office.com](https://outlook.office.com)
2. Open any email
3. Click **"..."** (More actions) → **"Get Add-ins"**
4. Click **"My add-ins"** → **"Add a custom add-in"** → **"Add from file"**
5. Upload your updated `manifest.xml`

### For Outlook Desktop:

1. **Windows**: 
   - File → Options → Trust Center → Trust Center Settings
   - Trusted Add-in Catalogs
   - Add `https://your-project-name.vercel.app` as a catalog
   - Restart Outlook
   - Home → Get Add-ins → My Organization

2. **Mac**:
   - Tools → Get Add-ins
   - My add-ins → Add from file
   - Select `manifest.xml`

---

## Troubleshooting

### Add-in doesn't load
- Verify manifest.xml URLs are correct
- Check browser console for errors
- Ensure Vercel deployment is live

### Authentication fails
- Verify Azure redirect URI matches exactly
- Check browser console for MSAL errors
- Ensure popup blockers are disabled

### Backend API calls fail
- Verify CORS is configured correctly on backend
- Check backend URL in config.js
- Test backend URL directly in browser: `https://your-backend.railway.app/api/prs/open`
- Check Railway/Render logs for errors

### CORS errors
- Add your Vercel domain to backend CORS allowed origins
- Ensure backend allows credentials if needed

---

## Quick Reference

### Frontend (Vercel)
- **URL**: `https://your-project-name.vercel.app`
- **Update**: Push to GitHub or run `vercel --prod`
- **Config**: `manifest.xml`, `auth.js`, `config.js`

### Backend (Railway/Render)
- **URL**: `https://your-app.up.railway.app` or `https://your-app.onrender.com`
- **Update**: Push to GitHub (auto-deploys)
- **Config**: CORS settings, environment variables

### Azure
- **App**: HEXA-OUTLOOK-AGENT
- **Update**: Authentication → Redirect URIs

---

## Cost

- **Vercel**: Free tier (generous for static sites)
- **Railway**: Free tier ($5 credit/month)
- **Render**: Free tier (spins down after inactivity, but free)

All services have paid tiers if you need more resources.

---

## Next Steps

1. ✅ Deploy frontend to Vercel
2. ✅ Deploy backend to Railway/Render
3. ✅ Update all URLs in code
4. ✅ Update Azure redirect URI
5. ✅ Redeploy frontend
6. ✅ Sideload updated manifest.xml
7. ✅ Test from different device!
