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
- [ ] Wait for GitHub Actions workflows to complete (~10-15 minutes)

## Automated Release Verification

All publishing is automated. Verify these workflows complete successfully:

### Core Release (`.github/workflows/release.yml`)
- [ ] GitHub Release created at https://github.com/bgreenwell/xleak/releases/tag/vX.Y.Z
- [ ] All artifacts present (binaries, tarballs, installers, checksums)
- [ ] Homebrew formula published to [homebrew-tap](https://github.com/bgreenwell/homebrew-tap)
- [ ] Published to [crates.io](https://crates.io/crates/xleak)

### Scoop Publishing (`.github/workflows/publish-scoop.yml`)
- [ ] Manifest updated in [scoop-bucket](https://github.com/bgreenwell/scoop-bucket)
- [ ] Manifest uses correct ZIP file and SHA256

### AUR Publishing (`.github/workflows/publish-aur.yml`)
- [ ] PKGBUILD updated in [AUR package](https://aur.archlinux.org/packages/xleak-bin)
- [ ] .SRCINFO generated correctly
- [ ] Version, URL, and SHA256 correct

### WinGet Publishing (`.github/workflows/publish-winget.yml`)
- [ ] PR created to [microsoft/winget-pkgs](https://github.com/microsoft/winget-pkgs)
- [ ] PR validation checks passing
- [ ] PR merged (may take 1-2 days, requires Microsoft approval)

## Test Installations

Test at least one platform from each category:

- [ ] Homebrew (macOS/Linux): `brew upgrade xleak && xleak --version`
- [ ] Scoop (Windows): `scoop update xleak && xleak --version`
- [ ] AUR (Arch Linux): `yay -Syu xleak-bin && xleak --version`
- [ ] WinGet (Windows): `winget upgrade bgreenwell.xleak` (after PR merge)
- [ ] Shell installer: Test install script from releases
- [ ] MSI installer: Download and test from GitHub Releases

## Troubleshooting

If any automated workflow fails:

1. Check GitHub Actions logs for specific errors
2. Verify secrets are configured:
   - `HOMEBREW_TAP_TOKEN`
   - `SCOOP_BUCKET_TOKEN`
   - `CARGO_REGISTRY_TOKEN`
   - `AUR_SSH_PRIVATE_KEY`
   - `WINGET_TOKEN`
3. See RELEASE_CHECKLIST.md for detailed troubleshooting
4. Fall back to manual publishing if needed (instructions in RELEASE_CHECKLIST.md)

### Known Issues

- **WinGet PR merge delay**: First-time submissions may take longer for review
- **AUR SSH timeout**: Retry workflow if AUR connection times out

## Post-Release

- [ ] All workflows completed successfully
- [ ] Installation tests passed on multiple platforms
- [ ] Announcement drafted (if applicable)
- [ ] Documentation updated if needed
- [ ] Monitor for installation issues on GitHub Discussions/Issues
- [ ] Close this issue

## Notes

<!-- Add any release-specific notes, blockers, or issues encountered -->

---

**Automation Status:**
- Fully automated: GitHub Releases, Homebrew, Scoop, crates.io, AUR
- Semi-automated: WinGet (PR created automatically, merge requires approval)
