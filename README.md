# XtremeLabs Tools

A comprehensive Document Automation & Workflow Suite for XtremeLabs, built with Python and Flask.

## Features

1. **Word Converter**: Convert Markdown files to beautiful Word documents with customizable watermarks, borders, formatting, and XtremeLabs branding.
2. **Feature Merger**: Insert a single Feature page into multiple target Word documents at once and automatically download them as a convenient ZIP archive.
3. **Practice Exam Tool** *(Coming Soon)*
4. **TOC Extractor** *(Coming Soon)*
5. **Scoring Guider** *(Coming Soon)*

---

## Local Development Setup

1. **Clone/Download the repository**
2. **Install local dependencies:**
   ```bash
   pip install -r requirements.txt
   ```
3. **Run the Flask Development Server:**
   ```bash
   python server.py
   ```
4. **Access the application:**
   Open your browser and navigate to `http://localhost:5000`

---

## Deployment to Render.com (Free Tier)

This project is fully ready to be deployed on Render's free tier. Follow these exact steps to host it live on the internet for free.

### Step 1: Upload Code to GitHub
1. Create a free account on [GitHub](https://github.com/).
2. Create a new "Public" repository.
3. Upload all the files from this folder (`app.py`, `server.py`, `converter.py`, `requirements.txt`, `templates/`, `static/`, etc.) to your new GitHub repository.

### Step 2: Deploy on Render
1. Go to [Render.com](https://render.com/) and create a free account (you can sign up with GitHub).
2. Once logged in, click "New" -> "Web Service".
3. Connect your GitHub repository that you created in Step 1.
4. Render will ask for some deployment settings. Fill them out exactly like this:
   - **Name:** xtremelabs-tools (or whatever you prefer)
   - **Environment:** `Python`
   - **Region:** `Frankfurt` (or closest to you)
   - **Branch:** `main` (or `master`)
   - **Build Command:** `pip install -r requirements.txt`
   - **Start Command:** `gunicorn app:app`
   - **Instance Type:** `Free`

### Step 3: Wait for Deployment
- Click **"Create Web Service"**.
- Render will start building your app. This might take 2-4 minutes.
- Once it says **"Live"**, you will see a URL at the top left (e.g., `https://xtremelabs-tools.onrender.com`).
- Click that URL, and your website is now live!

> **Note on Free Tier:** Render spins down free web services after 15 minutes of inactivity. This means if no one visits the site for a while, the *next* person to visit might experience a 50-second delay while the server wakes up. Once awake, it runs at normal speed.
