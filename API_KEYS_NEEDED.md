# API Keys Setup Instructions

Your API keys are **NOT** in the code right now. You need to add them locally.

## Required: Add Your Real Keys Locally

### 1. Update gemini-config.js
Replace `YOUR_GEMINI_API_KEY_HERE` with your actual Gemini API key

### 2. Update supabase-config.js
Replace these with your actual Supabase credentials:
- `YOUR_SUPABASE_URL_HERE` → `https://ejynlffmwiscbxrfaiyk.supabase.co`
- `YOUR_SUPABASE_ANON_KEY_HERE` → Your Supabase anon key from dashboard

## These files are now gitignored and won't be committed

When you commit, these files won't be included because they're in .gitignore.
Other developers will need to create their own versions with their keys.
