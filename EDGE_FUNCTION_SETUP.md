# Supabase Edge Function Deployment Guide

## Overview
The Gemini API key is now secured in a Supabase Edge Function. Your frontend calls the Edge Function, which then forwards requests to Gemini with the secure API key.

## Prerequisites
- Supabase CLI installed: `npm install -g supabase`
- Your Supabase project linked

---

## Step 1: Install Supabase CLI

**Option A: Using Scoop (Recommended for Windows)**
```powershell
# Install Scoop if you don't have it
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
Invoke-RestMethod -Uri https://get.scoop.sh | Invoke-Expression

# Install Supabase CLI
scoop bucket add supabase https://github.com/supabase/scoop-bucket.git
scoop install supabase
```

**Option B: Using Chocolatey**
```powershell
choco install supabase
```

**Option C: Use NPX (No installation needed)**
```powershell
# Replace 'supabase' with 'npx supabase' in all commands below
npx supabase login
npx supabase link --project-ref YOUR_PROJECT_REF
# etc...
```

---

## Step 2: Login to Supabase

```powershell
supabase login
```

---

## Step 3: Link Your Project

```powershell
cd "c:\Users\Dell\Documents\SchedulePro"
supabase link --project-ref YOUR_PROJECT_REF
```

**To find your project ref:**
- Go to [Supabase Dashboard](https://supabase.com/dashboard)
- Select your project
- Settings → General → Reference ID

---

## Step 4: Set the Gemini API Key as a Secret

```powershell
supabase secrets set GEMINI_API_KEY=your_actual_gemini_api_key_here
```

**Get your Gemini API key from:** https://aistudio.google.com/apikey

---

## Step 5: Deploy the Edge Function

```powershell
supabase functions deploy gemini-proxy
```

---

## Step 6: Update Your Local Config Files

Since the frontend now calls the Edge Function (not Gemini directly), you still need Supabase credentials locally:

**In `supabase-config.js`:**
```javascript
const SUPABASE_URL = 'https://your-project.supabase.co';
const SUPABASE_PUBLISHABLE_KEY = 'your_supabase_anon_key';
```

**You NO LONGER NEED `gemini-config.js`** - the Gemini key is now secure in Supabase!

---

## Step 7: Test the Function

```powershell
# Test locally (optional)
supabase functions serve gemini-proxy

# Or test directly in your app after deployment
```

---

## What Changed?

### Before:
```
Browser → Gemini API (with exposed key)
```

### After:
```
Browser → Supabase Edge Function → Gemini API (key hidden)
```

---

## Troubleshooting

### Function not working?
```powershell
# Check function logs
supabase functions logs gemini-proxy

# Verify secret is set
supabase secrets list
```

### Need to update the key?
```powershell
supabase secrets set GEMINI_API_KEY=new_key_here
```

---

## Security Notes

✅ Gemini API key is now stored securely in Supabase (never in code)  
✅ Edge Function runs on Supabase's servers (not in browser)  
✅ Config files are gitignored  
✅ Users cannot see or steal your Gemini API key  

---

## Next Steps

1. Deploy the function (see Step 5 above)
2. Remove the old Gemini key from `gemini-config.js` (or delete the file)
3. Test your app - the AI chat should work as before
4. Commit and push your code safely!
