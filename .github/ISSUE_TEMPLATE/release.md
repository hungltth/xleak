---
name: Release
about: Track a new version release
title: 'Release vX.Y.Z'
labels: release
assignees: ''

---

## Release Version

**Version:** vX.Y.Z

## Pre-Release Checklist

- [ ] All tests passing: `cargo test`
- [ ] No clippy warnings: `cargo clippy`
- [ ] Code is formatted: `cargo fmt --check`
- [ ] CHANGELOG.md updated with new version changes
- [ ] Version bumped in `Cargo.toml`
- [ ] Test binary works: `cargo run --release -- test_data.xlsx`

## Create Release

- [ ] Commit version bump: `git commit -m "chore: release X.Y.Z"`
- [ ] Push to main: `git push`
- [ ] Create version tag: `git tag vX.Y.Z`
- [ ] Push tag: `git push origin vX.Y.Z`
- [ ] GitHub Actions workflow triggered and completed successfully

## Verify Automated Releases

- [ ] GitHub Release created at https://github.com/bgreenwell/xleak/releases/tag/vX.Y.Z
- [ ] All artifacts present in GitHub Release
- [ ] Homebrew formula published to [bgreenwell/homebrew-tap](https://github.com/bgreenwell/homebrew-tap)
- [ ] Scoop manifest published to [bgreenwell/scoop-bucket](https://github.com/bgreenwell/scoop-bucket)

## Manual: Publish to AUR

- [ ] Generate PKGBUILD: `cargo aur`
- [ ] Get SHA256 hash from release
- [ ] Update PKGBUILD with correct source URL and hash
- [ ] Copy to AUR repo: `cp target/cargo-aur/PKGBUILD ~/xleak-bin/`
- [ ] Generate .SRCINFO with Docker
- [ ] Commit and push to AUR master branch
- [ ] Verify on [AUR](https://aur.archlinux.org/packages/xleak-bin)

## Test Installations

Test on at least one platform from each category:

- [ ] Homebrew (macOS/Linux)
- [ ] Scoop (Windows)
- [ ] AUR (Arch Linux)
- [ ] Shell installer (Linux/macOS)
- [ ] MSI installer (Windows)

## Post-Release

- [ ] Announcement published (if applicable)
- [ ] Documentation updated if needed
- [ ] Close this issue

## Notes

<!-- Add any release-specific notes, issues encountered, or deviations from the standard process -->
