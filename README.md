# Intune Windows PowerShell Script Assignment Reporter (Group Include/Exclusion & Built-in Targets)

> Exports a tenant report of Intune Windows PowerShell script assignments, resolving built-in targets (All Devices/All Users) and included/excluded Entra groups for audit-ready change reviews.

## Overview

This PowerShell project was extracted from an article on [MSEndpoint.com](https://msendpoint.com) and is part of the professional automation portfolio.

### Quick Start

```powershell
# Clone the repository
git clone https://github.com/Msendpoint/intune-windows-ps-script-assignment-reporter.git
cd intune-windows-ps-script-assignment-reporter

# One-click install & run (checks prerequisites, installs modules, launches main script)
.\Install.ps1
```

## Tech Stack

PowerShell 7+, Microsoft Graph SDK v2, OAuth (M365 SaaS Engine skeleton via PHP endpoint)

## Requirements

- PowerShell 7.0 or later
- Microsoft Graph CLI (if applicable)
- Appropriate administrative permissions

## About the Author

**Souhaiel Morhag** — Microsoft 365 & Intune Automation Specialist

- 🌐 **Blog**: [https://msendpoint.com](https://msendpoint.com) — Enterprise endpoint management insights
- 📚 **Academy**: [MSEndpoint Academy](https://app.msendpoint.com/academy) — Professional training & certifications
- 💼 **LinkedIn**: [Souhaiel Morhag](https://linkedin.com/in/souhaiel-morhag) — Connect for collaboration
- 🐙 **GitHub**: [@Msendpoint](https://github.com/Msendpoint)

## Source Article

This project originated from:
**[Group exclusions and built-in group targeting for PowerShell scripts on Windows — Context: [Week of November 16, 2026 (Service release 2611)] [Device configuration] [Applies to: Windows] Microsoft Intune now supports group exclusions and the built-in All ](https://msendpoint.com/article/group-exclusions-and-built-in-group-targeting-for-powershell-scripts-on-windows-context-week-of-november-16-2026-service-release-2611-device-configuration-applies-to-windows-microsoft-intune-now-supports-group-exclusions-and-the-built-in-all-users-and-all)**

Published on MSEndpoint.com — Enterprise endpoint management blog

## License

MIT License — see [LICENSE](LICENSE) for details.