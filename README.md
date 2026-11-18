<div align="center">

# 🐘 SRI VINAYAKA YUVA SENA
### 🚩 Community Festival Savings & Contribution Tracker

**భక్తి • ఐక్యత • సేవ — Devotion • Unity • Service**

A modern, mobile-first community web application designed for transparent fund tracking, festival expense management, verified payment receipts, and instant WhatsApp communication for annual Ganesh Chaturthi celebrations.

[🌐 **Open Live Website**](https://srivinayakayuvasena.vercel.app/) · [📂 **GitHub Repository**](https://github.com/tharungutta/Sri_VinayakaYuvasena_2)

[![Live Website](https://img.shields.io/badge/Live%20Website-srivinayakayuvasena.vercel.app-ff6b81?style=for-the-badge&logo=vercel)](https://srivinayakayuvasena.vercel.app/)
[![Platform](https://img.shields.io/badge/Platform-Web%20%2F%20PWA-6c5ce7?style=for-the-badge)](https://srivinayakayuvasena.vercel.app/)
[![Hosting](https://img.shields.io/badge/Hosted%20on-Vercel-000000?style=for-the-badge&logo=vercel)](https://srivinayakayuvasena.vercel.app/)
[![Database](https://img.shields.io/badge/Database-Google%20Sheets%20Live-0f9d58?style=for-the-badge&logo=googlesheets)](https://srivinayakayuvasena.vercel.app/)
[![Storage](https://img.shields.io/badge/Storage-Google%20Drive-4285f4?style=for-the-badge&logo=googledrive)](https://srivinayakayuvasena.vercel.app/)

</div>

---

## 🌟 Overview

**SRI VINAYAKA YUVA SENA (SVYS)** tracker provides 100% financial transparency to community members. All monthly contributions, payment screenshots, and festival expenditures are synced live with Google Sheets and Google Drive, allowing any devotee or member to review collections in real time from any smartphone or computer.

---

## ✨ Key Features

### 1. 🌸 Logo Devotional Seva
- **Interactive Pill Buttons**: Compact `🌸 Offer Flowers` and `🔔 Ring Bell` buttons positioned right beneath Lord Ganesha's circular emblem.
- **Logo-Contained Flower Shower**: Falling petals are strictly contained within the circular logo frame with realistic random drift and auto-cleanup.
- **Temple Bell Acoustics**: Web Audio API synthesized brass bell acoustics with harmonics (880 Hz, 1760 Hz, 2640 Hz) and physical bell swing animation.

### 2. 🗓️ Modern Month Breakdown Modal
- **5 Comprehensive Metric Cards**:
  - 🎯 **Target / Member**: Minimum monthly contribution (e.g., ₹200).
  - 💰 **Total Collected**: Total funds collected for that month.
  - 📊 **Group Target**: Total expected collection for the entire group.
  - 👥 **Members Paid**: Live count of members who have completed their contribution.
  - ⭐ **Surplus / Extra**: Total surplus collected beyond the target.
- **Member Cards with Avatars**:
  - Circular initials avatar badge.
  - Member name and 10-digit mobile number.
  - Dual-tone progress bar with payment date timestamp.
  - Status pill (`PAID`, `PARTIAL`, `PENDING`).
  - `📸 View Proof` button opening the verified Google Drive payment screenshot.
- **In-Modal Search & Filters**: Live search by member name/phone with filter pills (`All`, `Paid`, `Pending`).

### 3. 🔐 Role-Based Admin Access & 1-Click Logout
- **Public View Mode**: Regular visitors only see public statistics, member lists, and verified screenshots. All action buttons (`🧾 Receipt`, `📢 Remind`, `Send WhatsApp Statement`, `Delete Expense`) are strictly omitted from the DOM and protected by CSS rule (`body:not(.is-admin)`).
- **Admin Authentication**: Authenticate using committee credentials (`Tharun` / `Tharun@18`).
- **1-Click Admin Logout**:
  - Top navigation bar: Red `[🚪 Logout]` button beside `[🛡️ Admin: Tharun]`.
  - Admin Center modal: Header button and dedicated `🚪 Log Out Admin` tab.
  - Month Breakdown modal: `[🛡️ Admin Active] [🚪 Logout]` pill.
  - Sidebar: `Sign Out Admin` button.
  - Clicking Logout immediately revokes admin access and hides all receipt/reminder action buttons.

### 4. 🧾 Concise WhatsApp Payment Receipts
- **Short & Professional**: Clean, single-copy message under 10 lines with no duplicate body.
- **Unicode Stability**: Uses universally supported bullet points (`•`) and stable symbols (`🚩`, `₹`, `🙏`) to prevent font corruption on mobile devices.
- **Anti-Duplication**: 800ms debounce protection prevents accidental double-sends from rapid clicks.
- **Direct Live Link**: Embeds `https://srivinayakayuvasena.vercel.app/`.

### 5. 📸 High-Contrast Screenshot Viewer
- Modern dark-mode lightbox modal with Google Drive thumbnail resolution enhancement (`sz=w1600`).
- Fallback card with direct `🔗 Open Receipt in Google Drive ↗` button for protected permissions.

### 6. ⚡ Dynamic UPI QR & Direct App Payments
- **Committee UPI**: `9390266268@upi` (*THAMMINENI REVANTH*).
- **QR Code Download**: Save official UPI QR code directly to local device with one click.
- **Direct App Launch**: 1-click launch for PhonePe, Google Pay, Paytm, and standard UPI intent.

### 7. 🕉️ Festival Expenses & Treasury Budget Tracker
- Tracks expenditures across 7 festival categories (Murti & Idol, Mandap & Lighting, Sound & DJ, Pooja & Flowers, Annadanam, Nimajjanam Procession, Crackers & Misc).
- Live Treasury balance comparison: `Total Collections - Total Expenses = Net Balance`.
- Export financial statements to Excel (`.xls`).

### 8. 📲 Progressive Web App (PWA) & Offline Cache
- Installable on Android & iOS Home Screens via `manifest.json`.
- Offline resilience powered by `live-data.js` and `sw.js` (Service Worker).

---

## 🛠️ Architecture & Technology Stack

| Layer | Technology | Details |
|---|---|---|
| **Frontend** | HTML5, CSS3, ES6+ JavaScript | Zero external build tools; pure native web technologies |
| **Styling** | Custom Properties (CSS variables) | Warm saffron, gold, and royal maroon temple festival palette |
| **Hosting** | Vercel | Automatic deployments on git push configured via `vercel.json` |
| **Live Database** | Google Sheets | Live spreadsheet storing Members, Payments, Expenses, and Login History |
| **File Storage** | Google Drive | Auto-categorized folders for payment screenshots and expense bills |
| **Backend API** | Google Apps Script (`apps-script.gs`) | Web App endpoint serving GET/POST JSON APIs |
| **Audio** | Web Audio API | Synthesized temple bell chimes |

---

## 🔑 Administrator Credentials

- **Username:** `Tharun`
- **Password:** `Tharun@18`

*Note: Administrator tools are private and designed for recording contributions, managing members, logging expenses, and issuing official WhatsApp receipts.*

---

## 🚀 Deployment Guide (Vercel)

This repository includes a pre-configured `vercel.json` in the root directory:

```json
{
  "$schema": "https://openapi.vercel.sh/vercel.json",
  "framework": null,
  "buildCommand": null,
  "outputDirectory": "SVYS_COMPLETE_PACKAGE",
  "cleanUrls": true
}
```

Whenever you push to the `main` branch of `https://github.com/tharungutta/Sri_VinayakaYuvasena_2`, Vercel automatically deploys the contents of `SVYS_COMPLETE_PACKAGE` to [https://srivinayakayuvasena.vercel.app/](https://srivinayakayuvasena.vercel.app/).

---

## 📁 Repository Structure

```
Sri_VinayakaYuvasena_2/
├── vercel.json                 # Vercel deployment configuration
├── README.md                   # Repository overview & documentation
└── SVYS_COMPLETE_PACKAGE/      # Production deployment package
    ├── index.html              # Main single-page web app
    ├── styles.css              # Complete responsive stylesheet & design system
    ├── config.js               # Committee UPI, Apps Script, and Firebase configurations
    ├── dashboard.js           # Analytics and charts
    ├── auth.js                 # Authentication helper utilities
    ├── auth.css                # Authentication gate styling
    ├── live-data.js            # Offline data fallback & snapshot cache
    ├── manifest.json           # PWA web app manifest
    ├── sw.js                   # Service Worker for offline support
    ├── apps-script.gs          # Google Apps Script Web App source code
    ├── README.md               # Package-level documentation
    └── assets/                 # App icons, QR codes, and imagery
        ├── logo.png            # Official circular emblem
        ├── upi-qr.png          # Verified committee UPI QR code
        ├── qrcode.min.js       # Lightweight offline QR code generator
        ├── phonepe-icon.png    # PhonePe branding icon
        └── paytm-icon.webp     # Paytm branding icon
```

---

<div align="center">

**గణపతి బప్పా మోరియా! 🌺 Ganpati Bappa Morya!**  
*Sri Vinayaka Yuva Sena Committee*

</div>
