# Google to M365 Migration Toolkit

Controlled, enterprise-safe migration of **Gmail, Calendar, and Contacts** from Google Workspace to Microsoft 365.

## Quick Start (auto-locating launchers)

The launchers automatically find the scripts — run them from any directory.

```powershell
# Inspect your Google Cloud environment (read-only)
.\Run-Bootstrap.ps1 -Mode Inspect

# Run a read-only migration preflight
.\Run-Migration.ps1 -Mode Preflight -CsvPath .\examples\migration-users.csv -SkipExchangeConnection
```

## Structure

```
google-to-m365-migration/
├── Run-Bootstrap.ps1                  # Auto-locating launcher → Google Cloud prep
├── Run-Migration.ps1                  # Auto-locating launcher → migration controller
├── scripts/
│   ├── GoogleCloudMigrationBootstrap.ps1   # Inspect, SDK, APIs, service account, keys
│   └── GoogleToM365Migration.ps1           # Preflight, endpoint, batch, monitor, complete
├── docs/
│   ├── SETUP-GUIDE.md                 # Step-by-step setup
│   └── TROUBLESHOOTING.md             # Common errors & fixes
├── examples/
│   └── migration-users.csv            # Sample CSV format
├── config/
│   └── .env.example                   # Environment variables template
└── LICENSE
```

## Documentation

- **[Setup Guide](docs/SETUP-GUIDE.md)** — Full prerequisites and step-by-step setup
- **[Troubleshooting](docs/TROUBLESHOOTING.md)** — Common errors and fixes

## Author

**Arwaz Khan** — [arwazitbp2003@gmail.com](mailto:arwazitbp2003@gmail.com)
