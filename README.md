## Run on localhost

Quick start (Windows):

1) Double-click `serve.bat` in this folder
   - This starts a server at `http://localhost:5500`
   - If PowerShell asks about execution policy, choose "Run once" or allow

2) Or, in PowerShell:

```powershell
cd "C:\Users\SUNRAY\Desktop\BLO_APP"
.\serve.ps1 -Port 5500
```

If you prefer manual commands:
- With Python:
```powershell
cd "C:\Users\SUNRAY\Desktop\BLO_APP"
py -m http.server 5500
```
- With Node.js (no install):
```powershell
cd "C:\Users\SUNRAY\Desktop\BLO_APP"
npx --yes serve -l 5500 .
```

Then open: `http://localhost:5500`


## Deploy to Vercel

This is a static site. You can deploy it to Vercel in minutes.

1) Create a new GitHub repository and push this folder’s contents.

2) In Vercel:
   - Import your GitHub repo
   - Framework preset: “Other” (Static Site)
   - Output directory: root (leave blank)
   - Click Deploy

Notes:
- `vercel.json` is included to set helpful headers so the service worker (`sw.js`) and `manifest.webmanifest` update properly.
- The app is a PWA. After the first load over HTTPS, you can install it using the bottom banner.

### Optional: Vercel CLI
```bash
npm i -g vercel
vercel --prod
```

