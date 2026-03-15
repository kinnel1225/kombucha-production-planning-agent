#!/bin/bash
# ============================================================
#  KOMBUCHA BREWERY — GITHUB DEPLOYMENT GUIDE
#  Run these commands in your terminal to publish the project
# ============================================================

# ── STEP 1: Create the GitHub repo
# Go to https://github.com/new
# Repo name:     kombucha-production-planner
# Description:   Dynamic Production Planning AI Agent for a Kombucha Brewery | IE Portfolio
# Visibility:    Public  ← required for free GitHub Pages
# ✅ Add a README: NO  (we already have one)
# Click "Create repository"

# ── STEP 2: Open your terminal and navigate to the project folder
cd path/to/kombucha_github   # change this to where you downloaded the files

# ── STEP 3: Initialize git and connect to GitHub
git init
git add .
git commit -m "🍵 Initial commit — Kombucha Production Planning AI Agent"
git branch -M main
git remote add origin https://github.com/YOUR-USERNAME/kombucha-production-planner.git
git push -u origin main

# ── STEP 4: Enable GitHub Pages
# 1. Go to your repo on GitHub
# 2. Click "Settings" (top menu)
# 3. Click "Pages" (left sidebar)
# 4. Under "Source" select: Deploy from a branch
# 5. Branch: main   Folder: /docs
# 6. Click Save
# 7. Wait ~60 seconds, then visit:
#    https://YOUR-USERNAME.github.io/kombucha-production-planner

# ── STEP 5: Update the README and app with your username
# In README.md, replace all instances of:
#   YOUR-USERNAME  →  your actual GitHub username
# In docs/index.html, replace:
#   YOUR-USERNAME  →  your actual GitHub username

git add .
git commit -m "📝 Update GitHub username in README and web app"
git push

# ── FUTURE UPDATES: Push changes anytime
git add .
git commit -m "📅 Update production plan — March 2026"
git push
# GitHub Pages auto-deploys within ~30 seconds of each push

echo "Done! Your repo will be live at:"
echo "https://YOUR-USERNAME.github.io/kombucha-production-planner"
