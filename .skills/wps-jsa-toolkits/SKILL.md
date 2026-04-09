---
name: wps-jsa-toolkits
description: Maintain this WPS JSA macro utility repository. Use when updating project files such as fileSystemUtils.js or fileSystemUtils.md, adding WPS macro helper functions, renaming public utility modules, documenting APIs, validating JSDoc, or ensuring WPS/JSA built-ins are preferred before Windows ActiveX fallbacks.
---

# WPS JSA Toolkits

## Overview

Use this skill to maintain the WPS JSA macro helper library in this repository. Keep implementation, Chinese JSDoc, and Markdown reference documentation synchronized.

## Core Workflow

1. Inspect current files before editing: `fileSystemUtils.js`, `fileSystemUtils.md`, and any existing references to renamed modules.
2. Prefer existing public APIs and naming patterns. Avoid adding duplicate function names.
3. For file and folder logic, prefer WPS/JSA macro environment functions first, then use Windows objects only as fallback.
4. Add or update Chinese JSDoc for every function in `fileSystemUtils.js`, including `@param` and `@returns`.
5. Update `fileSystemUtils.md` whenever public behavior, parameters, return format, or examples change.
6. Run validation after edits: at minimum `node --check .\fileSystemUtils.js` and targeted searches for stale names or prefixes.

## WPS First

When implementing file system helpers:

- Prefer `Application.FileDialog` for user file or folder picking.
- Prefer `GetAttr`, `Dir`, `MkDir`, `FileLen`, `FileDateTime`, `FileCopy`, `Name`, `Kill`, and `RmDir` when they satisfy the behavior.
- Use `ActiveXObject("Scripting.FileSystemObject")`, `ActiveXObject("Shell.Application")`, or `ActiveXObject("WScript.Shell")` only after WPS/JSA-compatible approaches fail or when WPS/JSA has no equivalent recursive API.
- Keep failure behavior gentle: return `false`, `""`, `[]`, `{}`, or `-1` instead of throwing, unless the surrounding project establishes a different pattern.

## Documentation Rules

- Write JSDoc in Chinese.
- Include `@param` for each named parameter and `@returns` for every function.
- Document public functions in `fileSystemUtils.md` with purpose, parameters, return value, example code, and the example's goal.
- For object returns such as `getFilesByPath`, include the return shape in both JSDoc or the Markdown document when useful.
- Mark internal `_fsu...` helpers as implementation details and avoid recommending direct business macro calls to them.

## Project Agents

Use the bundled project agents as role guides when the task benefits from a narrower pass:

- [WPS JSA Macro Writer](agents/wps-jsa-macro-writer.md): use for implementing or updating WPS JSA macro utilities.
- [WPS JSA Macro Validator](agents/wps-jsa-macro-validator.md): use for checking changed macro code for WPS/JSA compatibility and platform runtime risk.

## Project References

Read [references/file-system-utils.md](references/file-system-utils.md) when working specifically on the file and folder helper module.



