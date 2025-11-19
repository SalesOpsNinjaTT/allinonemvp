# Setup Guide - All in One Dashboard

This guide will help you set up your development environment with clasp and GitHub.

## Prerequisites

### 1. Check if clasp is installed

```bash
clasp --version
```

If not installed:

```bash
npm install -g @google/clasp
```

### 2. Check if you're logged into clasp

```bash
clasp login --status
```

If not logged in:

```bash
clasp login
```

This will open a browser window. Log in with your Google account that has access to Google Apps Script.

---

## Part 1: Set Up clasp with Google Apps Script

### Step 1: Create a new Apps Script project

From the `all-in-one-dashboard` directory:

```bash
cd "/Users/konstantingevorkov/Documents/CODING/ALL IN ONE/all-in-one-dashboard"
clasp create --type standalone --title "All in One Dashboard"
```

This will:
- Create a new standalone Apps Script project in your Google Drive
- Generate a `.clasp.json` file with your script ID
- Link your local code to the remote Apps Script project

### Step 2: Push your code to Apps Script

```bash
clasp push
```

This uploads all your local `.js` files to the Apps Script editor.

### Step 3: Verify it worked

```bash
clasp open
```

This opens the Apps Script project in your browser. You should see:
- `Main.js`
- `ConfigManager.js`
- `HubSpotClient.js`
- `PipelineReview.js`
- `BonusCalculation.js`
- `EnrollmentTracker.js`
- `OperationalMetrics.js`

---

## Part 2: Set Up GitHub Repository

### Step 1: Create a GitHub repository

Go to [GitHub](https://github.com/new) and create a new repository:
- Repository name: `all-in-one-dashboard` (or your preferred name)
- Visibility: Private (recommended - contains business logic)
- **Don't** initialize with README (we already have one)

### Step 2: Add GitHub as remote

Replace `YOUR_USERNAME` with your GitHub username:

```bash
cd "/Users/konstantingevorkov/Documents/CODING/ALL IN ONE/all-in-one-dashboard"
git remote add origin https://github.com/YOUR_USERNAME/all-in-one-dashboard.git
```

### Step 3: Push to GitHub

```bash
git branch -M main
git push -u origin main
```

Enter your GitHub credentials when prompted.

---

## Part 3: Configure Apps Script Project

### Step 1: Set up Script Properties

1. Open your Apps Script project:
   ```bash
   clasp open
   ```

2. In the Apps Script editor:
   - Click **Project Settings** (gear icon)
   - Scroll to **Script Properties**
   - Click **Add script property**
   - Add these properties:

   | Property | Value |
   |----------|-------|
   | `HUBSPOT_ACCESS_TOKEN` | Your HubSpot API token |

### Step 2: Test the setup

1. In Apps Script editor, select `setupProject` from the function dropdown
2. Click **Run**
3. Approve permissions when prompted
4. Check the logs (Ctrl+Enter or View > Logs)

---

## Part 4: Development Workflow

### Daily Workflow

1. **Start of day - Pull latest**:
   ```bash
   clasp pull
   ```

2. **Make changes** in your local files (in `src/` directory)

3. **Push to Apps Script**:
   ```bash
   clasp push
   ```

4. **Test in Apps Script**:
   ```bash
   clasp open
   ```
   Run your functions and check logs

5. **View logs** (optional):
   ```bash
   clasp logs
   ```

6. **Commit to Git** when stable:
   ```bash
   git add -A
   git commit -m "Description of changes"
   git push
   ```

### Important: Always work locally

- ❌ **Never edit directly in Apps Script web editor**
- ✅ **Always edit in Cursor/VS Code locally**
- ✅ **Push with `clasp push`**
- ✅ **Commit to Git after testing**

---

## Part 5: Verify Everything Works

### Checklist

- [ ] clasp installed and logged in
- [ ] Apps Script project created (`clasp create`)
- [ ] Code pushed to Apps Script (`clasp push`)
- [ ] Can open project in browser (`clasp open`)
- [ ] GitHub repository created
- [ ] Code pushed to GitHub (`git push`)
- [ ] HubSpot API token configured in Script Properties
- [ ] Can run `setupProject` function without errors

---

## Troubleshooting

### clasp: command not found

```bash
npm install -g @google/clasp
```

### Permission denied when pushing

Run `clasp push --force` to overwrite remote files.

### .clasp.json not found

Make sure you're in the project directory:

```bash
cd "/Users/konstantingevorkov/Documents/CODING/ALL IN ONE/all-in-one-dashboard"
```

### GitHub authentication fails

Use a Personal Access Token instead of password:
1. Go to GitHub Settings > Developer settings > Personal access tokens
2. Generate new token with `repo` scope
3. Use token as password when pushing

---

## Next Steps

Once setup is complete:

1. Review [PRINCIPLES.md](../../PRINCIPLES.md) for development guidelines
2. Review [ALL_IN_ONE_DASHBOARD_ROADMAP.md](../../ALL_IN_ONE_DASHBOARD_ROADMAP.md) for the project plan
3. Start Phase 1: Extract HubSpot API client from VERSION 2.9

---

## Quick Reference

```bash
# clasp commands
clasp login                 # Login to Google account
clasp create               # Create new Apps Script project
clasp push                 # Push local code to Apps Script
clasp pull                 # Pull Apps Script code to local
clasp open                 # Open project in browser
clasp logs                 # View execution logs
clasp status               # Check what will be pushed

# Git commands
git status                 # Check what changed
git add -A                 # Stage all changes
git commit -m "message"    # Commit with message
git push                   # Push to GitHub
git pull                   # Pull from GitHub
```

---

**Need help?** Check the [clasp documentation](https://github.com/google/clasp) or [Git documentation](https://git-scm.com/doc).

