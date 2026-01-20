# DLL Analyzer

A .NET global tool to analyze DLL files and extract detailed information.

## Installation

### Step 1: Setup GitHub Packages Authentication

```bash
dotnet nuget add source https://github.com/tolaptrinhdh61-spec/DllAnalyzer/index.json \
  --name github \
  --username YOUR-GITHUB-USERNAME \
  --password YOUR-GITHUB-PAT \
  --store-password-in-clear-text
```

**Note:** You need a GitHub Personal Access Token with `read:packages` permission.

### Step 2: Install the tool

```bash
dotnet tool install --global dh.cs.dllanalyzer
```

### Step 3: Update (when new version available)

```bash
dotnet tool update --global dh.cs.dllanalyzer
```

## Usage

```bash
dllanalyzer
```

Place your DLL file as `DH.BLLCLS.dll` in the current directory.

## Output Files

- `{dll}.json` - Full analysis
- `{dll}.summary.json` - Summary
- `{dll}.external.json` - External references

## Uninstall

```bash
dotnet tool uninstall --global dh.cs.dllanalyzer
```
