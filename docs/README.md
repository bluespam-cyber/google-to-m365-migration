# Google to M365 Migration Toolkit

Controlled, enterprise-safe migration of **Gmail, Calendar, and Contacts** from Google Workspace to Microsoft 365 — a single PowerShell script, no Google Cloud SDK required.

## Quick Start (auto-locating launcher)

The launcher automatically finds the script — run it from any directory.

```powershell
# Print the Google Cloud + M365 setup steps (5-minute one-time prep)
.\Run-Migration.ps1 -Mode Guide

# Run a read-only migration preflight
.\Run-Migration.ps1 -Mode Preflight -CsvPath .\examples\migration-users.csv -KeyPath C:\keys\gws.json -GoogleAdminEmail admin@contoso.com -TargetDeliveryDomain o365.contoso.com

# Interactive wizard (auto-detects CSV + key, asks only what it can't find)
.\Run-Migration.ps1
```

## Structure

```
google-to-m365-migration/
├── Run-Migration.ps1                  # Auto-locating launcher → migration controller
├── scripts/
│   └── GoogleToM365Migration.ps1      # The entire migration controller (all 8 modes)
├── docs/
│   ├── SETUP-GUIDE.md                 # Step-by-step setup
│   └── TROUBLESHOOTING.md             # Common errors & fixes
├── examples/
│   └── migration-users.csv            # Sample CSV format
├── config/
│   └── .env.example                   # Parameter reference sheet
└── LICENSE
```

## Modes

| Mode | Description | Read-Only? |
|---|---|---|
| `Guide` | Prints the Google Cloud + M365 setup steps | ✅ Yes |
| `Preflight` | Validates CSV, routing domain, recipients, Google connectivity | ✅ Yes |
| `CreateEndpoint` | Creates the Gmail migration endpoint (shreds the key afterwards) | ❌ Mutation |
| `CreateBatch` | Creates the migration batch (never auto-starts it) | ❌ Mutation |
| `StartBatch` | Starts an existing batch | ❌ Mutation |
| `Monitor` | Monitors batch progress, optional per-user statistics | ✅ Yes |
| `Complete` | Completes the batch (cutover) | ❌ Requires `-ApproveCutover` |
| `Run` | End-to-end: endpoint → batch → start in one session | ❌ Mutation |

## Documentation

- **[Setup Guide](docs/SETUP-GUIDE.md)** — Full prerequisites and step-by-step setup
- **[Troubleshooting](docs/TROUBLESHOOTING.md)** — Common errors and fixes

## Author

**Arwaz Khan** — [arwazitbp2003@gmail.com](mailto:arwazitbp2003@gmail.com)