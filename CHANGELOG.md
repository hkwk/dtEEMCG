# Changelog

## v0.2.0 - 2025-12-28

### Added
- Cross-platform release workflow that builds on Linux/macOS/Windows and uploads platform-specific artifacts.
- Automated release creation on tag push and artifact upload.
- CI workflow ensures `cargo fmt`, `cargo clippy`, and `cargo test` run on push/PR.
- End-to-end unit test added for `dtEEMCG` processing logic.
- Capture and upload build logs for easier debugging of release builds.

### Changed
- Moved code to package `dttools` with two binaries (`dtEEMCG`, `dtproton`).
- Updated Excel processing logic (sheet renames, replacements, -999 handling, parentheses removal with red fill).

### Fixed
- Tests pass and workflows added to run CI and releases.


(See Git history for details)
