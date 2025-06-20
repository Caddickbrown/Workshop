# Computer Setup Automation Guide

A comprehensive checklist and automation guide for setting up new development environments. This standardizes the setup process across different computers and job roles, reducing time and ensuring consistency.

> **Note**: An initialization script is being developed to automate much of this process. See [Automation Status](#automation-status) below.

## 📋 Prerequisites

### System Requirements
- Windows 10/11 (primary target)
- Administrator privileges required for software installation
- Internet connection for downloading applications
- Minimum 8GB RAM recommended for development tools

### Permissions Needed
- **Administrator rights** for software installation
- **Excel Developer access** for macro installation
- **Registry modification rights** for some customizations
- **Network access** for cloud service connections

---

## 📦 Installation Checklist

### Core Applications
1. [Obsidian](https://obsidian.md/) - Note-taking and knowledge management
2. [VS Code](https://code.visualstudio.com/) - Primary code editor
3. [SQL Server Management Studio](https://docs.microsoft.com/en-us/sql/ssms/) - Database management

### Startup Folder Configuration
Location: `%APPDATA%\Microsoft\Windows\Start Menu\Programs\Startup`

Required files:
1. **KeepAwake.vbs** (Attached) - Prevents system sleep during work hours
2. **Outlook Shortcut** - Quick access to email
3. **ERP System Shortcut** - Direct access to work system

### Essential Downloads
1. **Scripts Folder** - [VBA Scripts Collection](https://github.com/Caddickbrown/Macros/blob/18f66736556730a727f62e3dd75afe92c00c6479/VBScripts)
2. **Excel Macros** - [Standard Installation Package](https://github.com/Caddickbrown/Macros/blob/18f66736556730a727f62e3dd75afe92c00c6479/VBA/Excel/Guidebook/Standard-Install.vba)

---

## ⚙️ Configuration Settings

### Excel Setup
1. **Enable Developer Tab**
   - File → Options → Customize Ribbon → Developer ✓
2. **Security Settings** 
   - File → Options → Trust Center → Trust Center Settings
   - Enable all macros (use with caution)
   - Trust access to VBA project object model ✓
3. **Quick Access Toolbar**
   - Add Filter button
   - Add Clear Filter button  
   - Add Macros button
   - Save configuration

### Chrome Browser Setup
1. **Account Setup**
   - Login/Create Google Account
   - Enable sync for bookmarks and extensions

2. **Essential Extensions**
   - [Speedtest by Ookla](https://chrome.google.com/webstore/detail/speedtest-by-ookla/pgjjikdiikihdfpoppgaidccahalehjh) - Network diagnostics
   - [QR Generator](https://chrome.google.com/webstore/detail/the-qr-code-extension/oijdcdmnjjgnnhgljmhkjlablaejfeeb) - Quick QR code creation
   - [Vimeo Record](https://chrome.google.com/webstore/detail/vimeo-record-screen-webca/ejfmffkmeigkphomnpabpdabfddeadcb) - Screen recording
   - [Video Downloader for Vimeo](https://chrome.google.com/webstore/detail/video-downloader-for-vime/cgmcdpfpkoildicgacgldinemhgmcbgp) - Video utilities
   - [Adblock](https://chrome.google.com/webstore/detail/adblock-%E2%80%94-best-ad-blocker/gighmmpiobklfepjocnamgkkbiglidom) - Ad blocking

### Taskbar Customization
Pin the following applications for quick access:
1. **Outlook** - Email and calendar
2. **Teams** (or equivalent communication tool)
3. **File Explorer** - File management
4. **Calculator** - Quick calculations
5. **Notepad** - Simple text editing
6. **Chrome** - Web browsing
7. **Excel** - Spreadsheet work
8. **ERP System** - Work-specific application

---

## 🤖 Automation Status

### Current Automation Level
- ✅ **Checklist Available** - Manual process documented
- ⏳ **Script Development** - Initialization script in progress
- ❌ **Full Automation** - Not yet implemented

### Planned Automation Features
- **One-click software installation** using package managers
- **Automatic configuration import** for applications
- **Registry modifications** for system optimizations
- **Extension batch installation** for browsers
- **Profile synchronization** across devices

### Development Roadmap
1. **Phase 1**: PowerShell script for software installation
2. **Phase 2**: Configuration file imports and settings
3. **Phase 3**: Full environment replication
4. **Phase 4**: Role-based customization profiles

---

## 🔧 Troubleshooting

### Common Installation Issues

#### Software Installation Failures
- **Issue**: "Access denied" during installation
- **Solution**: Run installer as Administrator, disable antivirus temporarily

#### Macro Security Warnings  
- **Issue**: Excel blocks macro execution
- **Solution**: Lower macro security or add trusted locations in Excel settings

#### Extension Installation Problems
- **Issue**: Chrome extensions won't install
- **Solution**: Check Chrome version compatibility, clear browser cache

#### VBA Script Errors
- **Issue**: Scripts fail to run after installation
- **Solution**: Verify Developer tab is enabled, check macro security settings

### Alternative Software Options

| Primary Choice | Alternative | Notes |
|----------------|-------------|-------|
| VS Code | Visual Studio | For .NET development |
| Chrome | Edge/Firefox | Browser preference |
| Obsidian | Notion/OneNote | Note-taking alternatives |
| SSMS | Azure Data Studio | Cross-platform option |

### Performance Optimization Tips
- **Disable startup programs** not in the essential list
- **Run disk cleanup** after installation
- **Update Windows** before installing development tools
- **Configure antivirus exclusions** for development folders

---

## 🎯 Role-Specific Customizations

### Data Analyst Profile
**Additional Software:**
- Power BI Desktop
- R/RStudio
- Tableau (if licensed)

**Excel Add-ins:**
- Analysis ToolPak
- Power Query/Power Pivot

### Developer Profile  
**Additional Software:**
- Git for Windows
- Docker Desktop
- Node.js/npm
- Python

**VS Code Extensions:**
- GitLens
- Python extension
- Prettier formatter

### Manager Profile
**Focus Areas:**
- Communication tools priority
- Reporting dashboards
- Simplified macro setup

---

## 🔄 Maintenance & Updates

### Monthly Checks
- [ ] Update all installed software
- [ ] Review and clean startup programs
- [ ] Backup current configurations
- [ ] Test critical macros and scripts

### Quarterly Reviews
- [ ] Evaluate new tools and extensions
- [ ] Update automation scripts
- [ ] Review security settings
- [ ] Document any customizations

---

## 📞 Support & Resources

### Internal Resources
- **Macro Repository**: GitHub - Caddickbrown/Macros
- **Script Documentation**: Internal wiki (if available)
- **IT Support**: Contact for enterprise software issues

### External Resources
- **Microsoft Documentation**: For Office and Windows issues
- **Stack Overflow**: For development-related problems
- **Vendor Support**: For specific application issues

---

*Last Updated: December 2024*  
*Next Review: March 2025*  
*Automation Script Status: In Development*