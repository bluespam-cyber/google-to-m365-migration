# Google to M365 Migration Toolkit

Automated Google Workspace → Microsoft 365 email migration.

## Quick Start

```powershell
# No-touch script (auto-installs everything)
& "$HOME\Downloads\NoTouchGoogleToM365Migration.ps1"

# Or interactive script (4 user-mapping modes)
& "$HOME\Downloads\AutoGoogleToM365Migration.ps1"
```

## Structure

```
├── scripts/
│   ├── NoTouchGoogleToM365Migration.ps1   # Full auto-deploy + migrate
│   └── AutoGoogleToM365Migration.ps1      # Interactive with 4 mapping modes
├── docs/
│   ├── SETUP-GUIDE.md                     # Step-by-step setup
│   └── TROUBLESHOOTING.md                 # Common errors & fixes
├── examples/
│   └── migration-users.csv                # Sample CSV format
├── config/
│   └── .env.example                       # Environment variables template
└── LICENSE
```

## Documentation

- **[Setup Guide](docs/SETUP-GUIDE.md)** — Full prerequisites and step-by-step setup
- **[Troubleshooting](docs/TROUBLESHOOTING.md)** — Common errors and fixes

## Author

**Arwaz Khan** — [arwazitbp2003@gmail.com](mailto:arwazitbp2003@gmail.com)
