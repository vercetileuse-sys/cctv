# 📹 CCTV Surveillance Bot — Aura Jewels

Automated daily camera health check system for Hikvision DVR/NVR devices.
Checks 506 cameras across 23 DVRs in under 5 minutes — fully automated.

## 🚀 Quick Start

### Option A — Web Dashboard (Recommended)
```bash
cd cctv_dashboard
pip install flask requests openpyxl urllib3 pillow opencv-python
python app.py
```
Open: http://localhost:5000  
Login: **Admin** / **Auracctv#2024**

### Option B — Command Line Bot (Excel only)
```bash
cd cctv_bot
pip install -r requirements.txt
python cctv_bot.py
```

---

## 📁 Repository Structure

```
cctv-surveillance-bot/
├── cctv_bot/                    # Standalone Python bot
│   ├── cctv_bot.py              # Main bot script (v9)
│   ├── nvr_config.json          # DVR/NVR credentials & skip_channels
│   ├── requirements.txt         # Python dependencies
│   └── RUN_AS_ADMIN.bat         # Windows one-click launcher
│
├── cctv_dashboard/              # Full web dashboard
│   ├── app.py                   # Flask backend + SQLite + auth
│   ├── nvr_config.json          # DVR/NVR credentials
│   ├── ecosystem.config.js      # PM2 config for production
│   ├── START_DASHBOARD.bat      # Windows one-click launcher
│   ├── templates/
│   │   ├── index.html           # Main dashboard UI
│   │   └── login.html           # Login page
│   ├── static/                  # CSS/JS assets
│   ├── snapshots/               # Camera snapshots (auto-created per scan)
│   ├── db/                      # SQLite database (auto-created)
│   └── logs/                    # PM2 logs
│
└── docs/
    ├── CCTV_Bot_Presentation.pptx   # Full presentation with infographics
    └── DEPLOYMENT_GUIDE.txt         # Step-by-step deployment guide
```

---

## 🔧 How It Works

### 1. Connection
The bot connects to each DVR/NVR using **Hikvision ISAPI** over HTTP — no screen clicking, no iVMS-4200 needed.

```
Bot → HTTP Digest Auth → DVR/NVR → ISAPI Response
```

### 2. Per-Camera Checks

| Check | Method | API |
|-------|--------|-----|
| Date & Time | DVR System Time API | `GET /ISAPI/System/time` |
| Camera Names | Channel List API | `GET /ISAPI/System/Video/inputs/channels` |
| Live Snapshot | Streaming API | `GET /ISAPI/Streaming/channels/{ch}/picture` |
| View Clarity | OpenCV 6-metric analysis | Local image processing |
| Recording | ContentMgmt Search | `POST /ISAPI/ContentMgmt/search` |

### 3. Clarity Checks (6 Metrics)
1. **File Size** → `NO SNAPSHOT` if < 3KB
2. **Brightness** → `NO VIDEO` / `VERY DARK` / `OVEREXPOSED`
3. **Std Deviation** → `LENS BLOCKED` if image is uniform
4. **Laplacian Variance** → `BLURRY` (center region only)
5. **Channel Difference** → `NIGHT VISION` if RGB channels match
6. **Timestamp Pixels** → `NO TIMESTAMP` if no white text in top-left

### 4. Recording Check (3-method cascade)
1. **ContentMgmt Search** — checks actual recorded files for today
2. **Track + HDD Status** — track exists + HDD ok = recording active
3. **Enable Flag** — for old DVR firmware

---

## 📊 Web Dashboard Features

| Section | Description |
|---------|-------------|
| 🏠 Dashboard | Live stats + vertical DVR list with dropdown |
| 📡 DVR / NVR | All 23 devices with expandable camera list |
| 🎥 All Cameras | 506 cameras — filterable by status |
| ⚠️ Issues | Problems only — arrow nav stays in filter |
| 📅 Scan History | Every scan in SQLite — Excel export per day |
| 📋 Activity Log | Live scan output |

---

## 🚀 Production Deployment (PM2)

```bash
# Install once
npm install -g pm2 pm2-windows-startup

# Start
cd cctv_dashboard
pm2 start ecosystem.config.js

# Auto-start on Windows boot
pm2 save
pm2-startup install

# Useful commands
pm2 list                    # Check status
pm2 logs cctv-dashboard     # View logs
pm2 restart cctv-dashboard  # Restart
```

---

## ⚙️ Configuration

Edit `nvr_config.json` to add/remove/modify devices:

```json
{
  "dvr_nvr_list": [
    {
      "name": "AGDVR 1",
      "ip_address": "192.168.2.237",
      "port": 80,
      "username": "admin",
      "password": "p@ssword",
      "enabled": true,
      "skip_channels": [27, 28, 29, 30, 31, 32]
    }
  ]
}
```

**skip_channels**: Channel numbers with no physical camera — excluded from checks.

---

## 📋 Requirements

```
Python 3.8+
flask
requests
openpyxl
urllib3
pillow
opencv-python
```

Optional for production:
```
Node.js (for PM2)
pm2
pm2-windows-startup
```

---

## 🔐 Login

- **URL**: http://localhost:5000
- **Username**: Admin
- **Password**: Auracctv#2024

---

## 📖 Documentation

See `docs/` folder:
- `CCTV_Bot_Presentation.pptx` — Full presentation with infographics
- `DEPLOYMENT_GUIDE.txt` — Step-by-step Windows deployment guide

---

**Aura Jewels Pvt. Ltd. — CCTV Operations Team**
