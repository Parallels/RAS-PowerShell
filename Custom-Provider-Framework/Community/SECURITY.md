# Security Policy

This is a community project of sample scripts and a test harness for the
Parallels RAS Custom Provider Framework. It is not official Parallels software
and is not supported by Parallels or Alludo. See the disclaimer in the
[README](README.md).

## Reporting a vulnerability

Please do not open a public issue for security problems.

Use GitHub's private vulnerability reporting instead: go to the **Security** tab
of this repository and choose **Report a vulnerability**. That opens a private
advisory visible only to the maintainers.

When reporting, include:

- the affected script and version (commit or release tag),
- a description of the issue and its impact,
- steps to reproduce, and any logs or proof of concept (with secrets removed).

You will get an acknowledgement, and we will work on a fix or a documented
mitigation. Because these are samples, some hardening is intentionally left to
the user; the disclaimer and the shared-responsibility model in the README and
the official documentation still apply.

## Handling secrets

Never commit real credentials. Provider settings such as tokens, secrets and
passwords belong in your local, untracked configuration and should be marked as
**secure** variables in the RAS Console. If you believe a secret was committed,
rotate it immediately and report it through the channel above.
