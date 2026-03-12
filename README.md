# Windows First Aid Toolkit

A practical first-step Windows repair utility built with AutoIt.

[![Built with AutoIt](https://img.shields.io/badge/Built%20with-AutoIt-5C2D91?style=for-the-badge)](https://www.autoitscript.com/site/autoit/)
[![Platform](https://img.shields.io/badge/Platform-Windows-0078D6?style=for-the-badge)](#)
[![Type](https://img.shields.io/badge/Type-Repair%20Utility-F59E0B?style=for-the-badge)](#)
[![Console UI](https://img.shields.io/badge/UI-Console%20Output-111111?style=for-the-badge)](#)
[![Help Files](https://img.shields.io/badge/Help-TXT%20%26%20HTML-2563EB?style=for-the-badge)](#)
[![Open Source](https://img.shields.io/badge/Open%20Source-Yes-10B981?style=for-the-badge)](#)
[![GitHub Repo](https://img.shields.io/badge/GitHub-Repository-181717?style=for-the-badge&logo=github)](https://github.com/Gexos/Windows-First-Aid-Toolkit)
[![Blog](https://img.shields.io/badge/Blog-gexos.org-0A58CA?style=for-the-badge)](https://gexos.org)

## Overview

**Windows First Aid Toolkit** is a technician-friendly Windows troubleshooting tool designed to group common repair, cleanup, and service tasks into one clean interface.

Instead of opening Command Prompt and typing the same repair commands again and again, you can use a single GUI with checkbox-based repair options, menu-driven actions, and a console-style output pane.

The goal of the app is simple:

- save time
- reduce repetitive repair work
- keep first-response troubleshooting organized
- provide a cleaner workflow for technicians and power users

## What the app does

Windows First Aid Toolkit focuses on safe and common first-step repair actions such as:

- Flush DNS
- Reset Winsock
- Release and renew IP
- Clear user temp folder
- Clear system temp folder
- Clear browser cache for common browsers
- Run `sfc /scannow`
- Run `DISM /Online /Cleanup-Image /RestoreHealth`
- Restart Windows Update
- Restart BITS
- Restart Print Spooler
- Create a restore point before changes

## Features

### Repair and cleanup options

- Preparation section with restore point option
- Network repair actions
- Cleanup actions
- Windows repair actions
- Service restart actions

### Console output styling

The lower output pane is designed like a technician console and supports different line styles, including:

- command lines
- normal output
- warnings
- success messages
- section headers

## Menus

### Presets

- Recommended
- Network repair only
- Cleanup only
- System repairs only
- Service restarts only

### Actions

- Run Selected
- Check All
- Clear All
- Save Log Copy
- Open Logs Folder
- Clear Console Output
- Exit

### Help

- Open Help TXT
- Open Help HTML
- About

## Default behavior

When the app starts, only the **Create restore point first** checkbox is selected.

All other repair options are left unchecked by default.

This keeps the initial state safer and makes it easier to build a repair run intentionally.

## Console output

The old activity log area has been replaced by a larger console-style output pane.

This pane shows:

- commands before they run
- command output
- warnings
- success lines
- section markers

This gives the application a more practical technician feel while still keeping output readable.

## Usage

### Run from source

1. Open the `.au3` script in AutoIt SciTE
2. Run it as Administrator
3. Select the repair actions you want
4. Use the **Actions** menu and choose **Run Selected**

### Recommended workflow

A practical workflow is:

1. Start with a restore point
2. Run only the specific actions you actually need
3. Watch the console output
4. Save a log copy if needed
5. Reboot only if one of the repairs requires it

```

## Files to keep beside the script or EXE

For the Help menu items to work correctly, keep these files in the same folder as the script or compiled EXE:

- `Windows_First_Aid_Toolkit_Help.txt`
- `Windows_First_Aid_Toolkit_Help.html`

## Safety notes

This tool is meant for **safe first-step troubleshooting**, not deep recovery or destructive repair work.

Keep these things in mind:

- Run the tool as Administrator
- `SFC` and `DISM` can take a long time
- Some actions may require a reboot
- Browser cleanup works best when browsers are closed
- Some files may be locked and skipped
- Restore point creation depends on System Protection being enabled

## Why this project exists

Windows technicians often repeat the same first-pass repair steps on different machines.

This project exists to turn that repetitive work into a cleaner process with:

- fewer repeated commands
- better visibility
- faster execution
- more consistent troubleshooting

## Credits

- **Author:** Giorgos Xanthopoulos
- **Icon credits:** Hristos Kalaitzis

## Links

- **GitHub:** https://github.com/Gexos/Windows-First-Aid-Toolkit
- **Blog:** https://gexos.org


## Contributing

Suggestions, fixes, and improvements are welcome.

If you test the tool and find a bug, open an issue in the repository and include:

- what you clicked
- what happened
- any AutoIt error message
- and whether you were running the script or a compiled EXE

---

**Windows First Aid Toolkit** is built to make common Windows first-response repair work faster, cleaner, and more practical.
