# 📊 Automated Daily GMB Rating & Review Tracker (All Outlets)

Tracks Google Business Profile (GMB) **Live Ratings** and **Review Counts** for all store outlets every day at **11:25 AM IST**, appends today's date column to `Outlet name and Link.xlsx` (preserving all previous days' history), and automatically emails the updated Excel file directly to your inbox.

---

## 📁 4 Tracked Outlet Tabs

1. **`Bikanervala GGN`**: All 24 Gurgaon outlets (FS61, FS50, ... F159) with full September historical data.
2. **`Competitors`**: 35 Haldiram & Om Sweets competitor outlets in Gurgaon.
3. **`India BV`**: 78 Bikanervala outlets across Delhi, UP, Punjab, etc.
4. **`International BV`**: 19 International Bikanervala outlets (Dubai, Nepal, NZ, USA, UK, etc.).

---

## 🚀 How to Run or Deploy

### Option 1: Run Locally on Your PC
```bash
pip install -r requirements.txt
playwright install chromium
python scraper.py
python send_email.py
```

### Option 2: Upload to GitHub for Automated Daily Emails
1. Create a new GitHub repository (e.g. `gmb-all-outlets-tracker`).
2. Upload all files from this `New folder/`:
   * `.github/workflows/daily_gmb_update.yml`
   * `scraper.py`
   * `send_email.py`
   * `Outlet name and Link.xlsx`
   * `requirements.txt`
   * `.gitignore`
   * `README.md`
3. Add your GitHub Repository Secrets (**Settings** > **Secrets and variables** > **Actions**):
   * `EMAIL_SENDER`: Your Gmail address
   * `EMAIL_PASSWORD`: Your 16-character Gmail App Password
   * `EMAIL_RECEIVER`: Where you want to receive the daily email
4. It will run automatically every day at 12:30 AM IST (19:00 UTC), update the sheet with today's column, and email you the report so it's ready in your inbox when you arrive.
