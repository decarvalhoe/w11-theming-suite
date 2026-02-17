# Contributing to Windows 11 Theming Suite

Thank you for your interest in contributing! This document provides guidelines and information to make the contribution process smooth for everyone.

## Table of Contents

- [Getting Started](#getting-started)
- [Development Setup](#development-setup)
- [Branch Strategy](#branch-strategy)
- [How to Contribute](#how-to-contribute)
- [Code Standards](#code-standards)
- [Commit Messages](#commit-messages)
- [Pull Request Process](#pull-request-process)
- [Reporting Issues](#reporting-issues)

---

## Getting Started

1. **Fork** the repository on GitHub
2. **Clone** your fork locally:
   ```bash
   git clone https://github.com/YOUR-USERNAME/w11-theming-suite.git
   cd w11-theming-suite
   ```
3. **Add the upstream remote**:
   ```bash
   git remote add upstream https://github.com/decarvalhoe/w11-theming-suite.git
   ```
4. **Create a feature branch** from `DEV`:
   ```bash
   git checkout DEV
   git pull upstream DEV
   git checkout -b feat/your-feature-name
   ```

---

## Development Setup

### Requirements

- Windows 11 (Build 22621+, ideally 25H2)
- PowerShell 5.1+ (ships with Windows)
- Administrator privileges (for testing DLL injection)
- Visual Studio Build Tools with MSVC x64 (only for native DLL changes)

### Loading the Module

```powershell
# Import with force-reload during development
Import-Module .\w11-theming-suite.psd1 -Force

# Verify all commands are exported
(Get-Command -Module w11-theming-suite).Count  # Should be 51
```

### Building Native DLLs

Only needed if you modify C++ code in `native/ShellTAP/` or `native/TaskbarTAP/`:

```cmd
# Requires VS Build Tools with MSVC x64
cd native\ShellTAP
build.cmd

cd native\TaskbarTAP
build.cmd
```

Pre-built binaries in `native/bin/` are committed to the repo for convenience.

---

## Branch Strategy

| Branch | Purpose | Merges From |
|--------|---------|-------------|
| `DEV` | Active development (default) | Feature branches |
| `STAGE` | Pre-release testing | `DEV` |
| `PROD` | Stable releases | `STAGE` |

- **All PRs target `DEV`**
- Feature branches: `feat/description`
- Bug fixes: `fix/description`
- Documentation: `docs/description`

---

## How to Contribute

### Adding a New Transparency Target

1. Use `Invoke-ShellTAPInject` with discovery mode to find XAML elements:
   ```powershell
   Invoke-ShellTAPInject -TargetProcess YourProcess -TargetId discovery -Mode Transparent
   ```
2. Review the log file for element names and types
3. Create `Invoke-YourTargetDiscovery` and `Invoke-YourTargetTransparency` functions
4. Follow the pattern established in `Invoke-StartMenuTransparency`

### Adding a Theme Preset

1. Create a new JSON file in `config/presets/`
2. Validate against `config/schema.json`
3. Include at minimum: `meta`, `mode`, and at least one visual section
4. Test with `Install-W11Theme -ConfigPath your-preset.json`

### Modifying the Config Schema

1. Update `config/schema.json` with new properties
2. Update the ThemeOrchestrator's `Install-W11Theme` to handle the new properties
3. Add appropriate defaults and validation

---

## Code Standards

### PowerShell

- **Strict mode**: All modules use `Set-StrictMode -Version Latest`
- **Verb-Noun naming**: Follow PowerShell approved verbs (`Get-Verb` for the list)
- **Prefix**: All public functions use the `W11` prefix (e.g., `Set-W11WindowBackdrop`)
- **Comment-based help**: Every exported function must have `.SYNOPSIS`, `.DESCRIPTION`, `.PARAMETER`, and `.EXAMPLE` blocks
- **Error handling**: Use `try/catch` with `Write-Host '[ERROR]'` pattern for user-facing errors
- **Output formatting**: Use the established pattern:
  ```powershell
  Write-Host '[OK]    ' -ForegroundColor Green -NoNewline
  Write-Host "Success message here."

  Write-Host '[ERROR] ' -ForegroundColor Red -NoNewline
  Write-Host "Error message here."

  Write-Host '[INFO]  ' -ForegroundColor Cyan -NoNewline
  Write-Host "Informational message."
  ```

### C++ (Native DLLs)

- **Target**: x64 only, Windows 10/11 SDK
- **Style**: Win32 API conventions, `wchar_t` strings
- **Logging**: Write to log file path from `ShellTAPConfig` struct
- **Memory**: Use shared memory segments prefixed with `W11ThemeSuite_`
- **Error handling**: Log failures, don't crash the host process

### JSON (Theme Configs)

- Validate against `config/schema.json`
- Use camelCase for property names
- Include `meta.name`, `meta.version`, `meta.author` at minimum

---

## Commit Messages

Follow the [Conventional Commits](https://www.conventionalcommits.org/) specification:

```
type: short description

Longer description if needed.
```

### Types

| Type | Description |
|------|-------------|
| `feat` | New feature |
| `fix` | Bug fix |
| `docs` | Documentation only |
| `refactor` | Code change that neither fixes a bug nor adds a feature |
| `test` | Adding or updating tests |
| `build` | Changes to build system or native DLLs |
| `chore` | Maintenance tasks |

### Examples

```
feat: add Flyout transparency via ShellTAP injection
fix: shared memory init race condition in DllMain
docs: add troubleshooting section to README
refactor: extract common TAP injection logic into helper
```

---

## Pull Request Process

1. **Ensure your branch is up to date** with `DEV`:
   ```bash
   git fetch upstream
   git rebase upstream/DEV
   ```

2. **Test your changes**:
   - Module loads without errors (`Import-Module -Force`)
   - All existing commands still work
   - New commands appear in `Get-Command -Module w11-theming-suite`
   - If modifying native DLLs: test injection on a live Windows 11 system

3. **Create the PR** targeting `DEV` with:
   - Clear title following commit message conventions
   - Description of what changed and why
   - Test plan (what you tested, on which Windows build)

4. **Address review feedback** with new commits (don't force-push during review)

5. **After approval**, the PR will be squash-merged into `DEV`

---

## Reporting Issues

### Bug Reports

Include:
- Windows 11 build number (`winver`)
- PowerShell version (`$PSVersionTable.PSVersion`)
- Steps to reproduce
- Expected vs actual behavior
- Any error messages (full text)

### Feature Requests

Include:
- Description of the desired behavior
- Use case (why is this useful?)
- If it involves a new Windows UI element: the process name and window class

### Security Issues

If you discover a security vulnerability, **do not open a public issue**. Instead, contact the maintainers directly through GitHub's private security advisory feature.

---

## Questions?

Open a [Discussion](https://github.com/decarvalhoe/w11-theming-suite/discussions) on GitHub for general questions about the project or development.
