# Security Policy

## Supported Versions

Security fixes are accepted for the current `1.x` release line.

## Reporting A Vulnerability

If GitHub Security Advisories are enabled for the repository, use a private advisory to report vulnerabilities.

If private advisories are not available, open a GitHub issue with a minimal description and avoid including sensitive workbook contents, credentials, customer data, or private paths.

Please include:

- affected version
- operating system and Python version
- minimal reproduction steps
- expected impact
- whether the issue requires opening an untrusted workbook

## Security Model

`xlsx-mcp` is local-first automation. It can read and write files that the running process can access.

Important considerations:

- Only connect the MCP server to clients you trust.
- Be careful when giving agents access to directories containing sensitive workbooks.
- Do not open untrusted `.xlsx` files in privileged environments.
- Save operations are explicit, but callers can still overwrite files when they request it.
- The project does not execute VBA macros or intentionally run workbook code.

## Sensitive Data

Do not attach private workbooks to public issues. If a workbook is needed to reproduce a bug, reduce it to a minimal synthetic file first.
