# Backend Deployment Guide

This guide shows you how to deploy your FastAPI backend to Railway or Render.

## Prerequisites

Your backend should be a FastAPI application with:
- `main.py` (or similar) with your FastAPI app
- `requirements.txt` with all dependencies
- CORS middleware configured

---

## Option 1: Deploy to Railway (Recommended - Easiest)

### Step 1: Prepare Your Backend

1. **Ensure CORS is configured** in your FastAPI app:

```python
from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware

app = FastAPI()

# Add CORS middleware
app.add_middleware(
    CORSMiddleware,
    allow_origins=[
        "https://your-frontend.vercel.app",  # Your Vercel URL
        "https://localhost:3000",  # For local testing
        "http://localhost:3000",   # For local testing
    ],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)
```

2. **Create `Procfile`** (optional, Railway auto-detects):
```
web: uvicorn main:app --host 0.0.0.0 --port $PORT
```

3. **Ensure `requirements.txt` exists** with all dependencies:
```
fastapi
uvicorn[standard]
python-dotenv
# ... your other dependencies
```

### Step 2: Deploy to Railway

1. **Go to [railway.app](https://railway.app)** and sign in with GitHub

2. **Click "New Project"** → **"Deploy from GitHub repo"**

3. **Select your backend repository**

4. **Railway will automatically**:
   - Detect it's a Python app
   - Install dependencies from `requirements.txt`
   - Start your app

5. **Get your backend URL**:
   - Go to your project → **Settings** → **Domains**
   - Railway provides: `your-app.up.railway.app`
   - Or add a custom domain

6. **Set environment variables** (if needed):
   - Go to **Variables** tab
   - Add any secrets (API keys, database URLs, etc.)
   - Example: `OPENAI_API_KEY=sk-...`

7. **Update CORS** in your code with the actual Railway URL, then:
   - Push changes to GitHub
   - Railway will auto-redeploy

### Step 3: Test Your Backend

```bash
# Test the health endpoint (if you have one)
curl https://your-app.up.railway.app/api/prs/open

# Or test in browser
open https://your-app.up.railway.app/docs  # FastAPI docs
```

---

## Option 2: Deploy to Render

### Step 1: Prepare Your Backend

Same as Railway - ensure CORS and `requirements.txt` are set up.

### Step 2: Deploy to Render

1. **Go to [render.com](https://render.com)** and sign in

2. **Click "New +"** → **"Web Service"**

3. **Connect your GitHub repository** (select your backend repo)

4. **Configure the service**:
   - **Name**: `procurement-backend` (or your choice)
   - **Environment**: **Python 3**
   - **Region**: Choose closest to you
   - **Branch**: `main` (or your default branch)
   - **Root Directory**: (leave empty if root)
   - **Build Command**: `pip install -r requirements.txt`
   - **Start Command**: `uvicorn main:app --host 0.0.0.0 --port $PORT`

5. **Set environment variables** (if needed):
   - Scroll down to **Environment Variables**
   - Add any secrets

6. **Click "Create Web Service"**

7. **Wait for deployment** (first deploy takes ~5 minutes)

8. **Get your URL**: `https://your-app.onrender.com`

### Step 3: Update CORS and Redeploy

Update your CORS configuration with the Render URL, push to GitHub, and Render will auto-redeploy.

---

## Important: CORS Configuration

Your backend **MUST** allow requests from your Vercel frontend. Update your CORS config:

```python
from fastapi.middleware.cors import CORSMiddleware

app.add_middleware(
    CORSMiddleware,
    allow_origins=[
        "https://your-frontend.vercel.app",  # Your actual Vercel URL
        "https://localhost:3000",            # For local dev
        "http://localhost:3000",             # For local dev
    ],
    allow_credentials=True,
    allow_methods=["GET", "POST", "PUT", "DELETE", "OPTIONS"],
    allow_headers=["*"],
)
```

**After deployment**, replace `your-frontend.vercel.app` with your actual Vercel URL.

---

## Environment Variables

If your backend needs environment variables (API keys, database URLs, etc.):

### Railway:
1. Go to project → **Variables** tab
2. Click **+ New Variable**
3. Add key-value pairs
4. App will restart automatically

### Render:
1. Go to service → **Environment** tab
2. Click **Add Environment Variable**
3. Add key-value pairs
4. Save and redeploy

---

## Testing Your Deployed Backend

1. **Check FastAPI docs**: `https://your-backend.railway.app/docs`
2. **Test an endpoint**:
   ```bash
   curl https://your-backend.railway.app/api/prs/open
   ```
3. **Check logs**:
   - Railway: Project → **Deployments** → Click deployment → **View Logs**
   - Render: Service → **Logs** tab

---

## Common Issues

### CORS Errors
- **Problem**: Frontend can't call backend
- **Solution**: Add your Vercel URL to `allow_origins` in CORS config

### Port Issues
- **Problem**: App crashes on startup
- **Solution**: Use `$PORT` environment variable (Railway/Render set this automatically)
- **Fix**: `uvicorn main:app --host 0.0.0.0 --port $PORT`

### Dependencies Missing
- **Problem**: Import errors in logs
- **Solution**: Ensure all dependencies are in `requirements.txt`

### Environment Variables Not Working
- **Problem**: API keys not found
- **Solution**: Set them in Railway/Render dashboard, not in code

---

## Cost

- **Railway**: Free tier with $5 credit/month (enough for testing)
- **Render**: Free tier (spins down after 15 min inactivity, but free)

Both have paid tiers for production use.

---

## Next Steps

After backend is deployed:

1. ✅ Note your backend URL (e.g., `https://your-app.up.railway.app`)
2. ✅ Update frontend `config.js` with backend URL
3. ✅ Update backend CORS with frontend URL
4. ✅ Redeploy both
5. ✅ Test from your add-in!
