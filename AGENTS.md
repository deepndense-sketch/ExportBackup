# Release and update rules

- When committing a plugin update for release, increment its version without waiting for a reminder. Use a patch bump for fixes and small changes, and a minor bump for new features.
- Keep `version.json`, both extension versions in `CSXS/manifest.xml`, and the displayed release version in sync. Update release notes for the actual changes.
- Preserve automatic GitHub update checks and the visible new-version download/install action. Do not silently install updates or interrupt active exports or copies. Explain when reopening the panel or restarting Premiere is required.
- Run appropriate checks before publishing and report the version and commit. Commit and push only when authorized by the user; these rules alone do not authorize publication.
