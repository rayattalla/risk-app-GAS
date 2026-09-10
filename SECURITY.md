# Security Baseline — LAUSD GAS Starter

## Absolute Rules
1. **No secrets in source code.** Ever.
2. All API keys, tokens, and credentials live only in `PropertiesService` (Script Properties).
3. Never commit `.clasprc.json`, `.env*`, service-account files, or any file containing credentials.
4. Web apps must use `"access": "DOMAIN"` unless there is a documented exception.
5. Prefer `executeAs: "USER_ACCESSING"` for user-facing web apps.

## Library vs Shell
- Shared logic and configuration live in `LAUSDLib`.
- Shell projects stay thin and only contain project-specific config mirrors + entry points.
- Keys are read only from PropertiesService inside the library.

## Before every push
- Confirm no secrets with a quick search.
- Confirm `.claspignore` is present.
- Confirm you are pushing the correct scriptId.
- Register / update the project in the App Registry (Script ID, status, security review).