# fw-recorder

> Local firewall request tracker — a web form that saves to Excel. Zero installs, zero dependencies.

## Overview

A single-page web application for tracking firewall rule requests, backed by a PowerShell HTTP server and Microsoft Excel. Designed for IT admins and network engineers who need to manage firewall requests without external services or database setup.

## Features

- **Web-based form** — fill in request details in a clean dark-themed UI
- **Excel-backed storage** — all data saved to `.xlsx` on your Desktop, natively sortable/filterable in Excel
- **Search & filter** — search by ID, IP, port, requester; filter by status and priority
- **Edit & update** — change status, add notes, mark as closed
- **CSV export** — one-click export to CSV for reporting
- **Zero installs** — runs on Windows with only PowerShell (built-in) and your browser

## Tech Stack

| Layer | Technology |
|-------|-----------|
| Frontend | Vanilla HTML/CSS/JavaScript (SPA, no framework) |
| Backend | PowerShell (`System.Net.HttpListener`) |
| Storage | Microsoft Excel (COM automation) |
| Deployment | Self-hosted on Windows, local network only |

## Quick Start

### Prerequisites

- **Windows 10/11** with PowerShell 5.1+ (built-in)
- **Microsoft Excel** installed (for .xlsx file creation)
- A web browser (Chrome, Edge, Firefox, etc.)

### Setup

1. Download or clone this repository:
   ```powershell
   git clone https://github.com/cenku613-ai/fw-recorder.git
   cd fw-recorder
   ```

2. Run the server:
   ```powershell
   .\firewall-tracker.ps1
   ```
   *(Optional: specify a port: `.\firewall-tracker.ps1 -Port 9090`)*

3. Open your browser to: **`http://localhost:18080`**

4. The first time you submit a form, an Excel file `firewall-requests.xlsx` is created on your Desktop.

## Fields

| Field | Description | Example |
|-------|-------------|---------|
| Application | Application/service name | HTTP, MySQL, SSH |
| Requester | Who requested the rule | John Doe |
| Priority | Urgency level | High / Medium / Low |
| Status | Current state | Pending / Approved / Rejected / Implemented |
| Source IP | Source address | 10.0.1.0/24 |
| Dest IP | Destination address | 10.0.2.0/24 |
| Dest Port | Destination port | 443 |
| Protocol | Network protocol | TCP / UDP / ICMP / Any |
| Direction | Traffic direction | Inbound / Outbound / Both |
| Justification | Business reason | Why this rule is needed |
| Ticket Ref | Related ticket | JIRA-1234, IT-5678 |
| Notes | Additional notes | Implementation details |
| Date Submitted | Auto-filled | yyyy-MM-dd HH:mm |
| Date Closed | When implemented/rejected | yyyy-MM-dd |

## Usage

### Adding a Request

1. Go to the **New Request** tab
2. Fill in the form fields
3. Click **Save**

### Viewing & Searching

1. Go to the **All Requests** tab
2. Use the search bar to find by ID, IP, requester, or port
3. Filter by status or priority using the dropdowns

### Editing a Request

1. Click **Edit** on any row
2. Update the status, priority, or notes
3. Optionally set a close date
4. Click **Save Changes**

### Exporting Data

Click **Export CSV** to download all records as a `.csv` file (Excel-compatible with BOM for UTF-8).

## Architecture

```
┌─────────────────┐         ┌──────────────────┐         ┌──────────────┐
│   Browser       │ HTTP    │ firewall-tracker │ Excel   │ Desktop      │
│  (form.html)    │ ◄─────► │ .ps1 (server)    │ ◄─────► │ .xlsx file   │
│                 │ POST    │ (HttpListener)   │ COM API │               │
└─────────────────┘         └──────────────────┘         └──────────────┘
```

- **`firewall-tracker.ps1`** — PowerShell HTTP server using `System.Net.HttpListener`
  - Serves the HTML form
  - REST API (`/api/records` GET/POST, `/api/records/<key>` PUT/DELETE)
  - Excel COM automation for reading/writing `.xlsx`
- **`form.html`** — Single-page application (no framework, no CDN)
  - Form submission via `fetch()` API
  - Table rendering with search/filter
  - Edit modal for status updates

## Security Notes

- **Local only** — server binds to `localhost` only, not accessible from other machines
- **No sensitive data** — no passwords, API keys, or credentials stored
- **Excel file** — stored on Desktop, unprotected (store securely if it contains sensitive network topology info)
- **PAT exposure** — if sharing this repo, ensure GitHub tokens in `.env` are not committed

## Troubleshooting

| Problem | Solution |
|---------|----------|
| Excel file not created | Ensure Microsoft Excel is installed; check Desktop permissions |
| PowerShell blocks execution | Run `Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass` first |
| Port already in use | Specify a different port: `.\firewall-tracker.ps1 -Port 9090` |
| CORS errors | Not applicable — server and form serve from same origin |

## License

MIT
