# Release Checklist

Use this checklist when preparing a new release of xleak. You can also create a GitHub issue using the "Release" template to track progress.

## Pre-Release

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
- [ ] Wait for GitHub Actions workflow to complete (~5-10 minutes)

## Verify Automated Releases

- [ ] GitHub Release created with release notes
- [ ] All artifacts present in GitHub Release:
  - [ ] Binaries for all platforms (Linux, macOS, Windows)
  - [ ] Tarballs (.tar.xz, .tar.gz)
  - [ ] ZIP archive for Windows
  - [ ] MSI installer for Windows
  - [ ] Shell/PowerShell installer scripts
  - [ ] SHA256 checksum files
- [ ] Homebrew formula published to [bgreenwell/homebrew-tap](https://github.com/bgreenwell/homebrew-tap)
- [ ] Scoop manifest published to [bgreenwell/scoop-bucket](https://github.com/bgreenwell/scoop-bucket)
- [ ] Published to crates.io: https://crates.io/crates/xleak

## Manual: Publish to AUR

- [ ] Generate PKGBUILD: `cargo aur`
- [ ] Get SHA256 hash from release:
  ```bash
  RELEASE_URL="https://github.com/bgreenwell/xleak/releases/download/vX.Y.Z/xleak-x86_64-unknown-linux-gnu.tar.xz.sha256"
  SHA256=$(curl -sL "$RELEASE_URL" | cut -d' ' -f1)
  echo $SHA256
  ```
- [ ] Update PKGBUILD in `target/cargo-aur/`:
  - [ ] Update source URL to point to correct tarball
  - [ ] Update `sha256sums` with hash from above
  - [ ] Verify `pkgver` matches release version
- [ ] Copy to AUR repo: `cp target/cargo-aur/PKGBUILD ~/xleak-bin/`
- [ ] Generate .SRCINFO using Docker:
  ```bash
  docker run --rm -v ~/xleak-bin:/build archlinux:latest /bin/bash -c \
    "useradd -m builder && cd /build && chown -R builder:builder . && \
     su builder -c 'makepkg --printsrcinfo' > .SRCINFO"
  ```
- [ ] Commit and push to AUR:
  ```bash
  cd ~/xleak-bin
  git add PKGBUILD .SRCINFO
  git commit -m "Update to vX.Y.Z"
  git push origin master
  ```
- [ ] Verify package appears on [AUR web interface](https://aur.archlinux.org/packages/xleak-bin)

## Test Installations

- [ ] **Homebrew (macOS/Linux)**:
  ```bash
  brew update
  brew upgrade xleak
  xleak --version
  ```

- [ ] **Scoop (Windows)**:
  ```powershell
  scoop update
  scoop update xleak
  xleak --version
  ```

- [ ] **AUR (Arch Linux)**:
  ```bash
  yay -Syu xleak-bin
  xleak --version
  ```

- [ ] **Shell installer (Linux/macOS)**:
  ```bash
  curl --proto '=https' --tlsv1.2 -LsSf \
    https://github.com/bgreenwell/xleak/releases/latest/download/xleak-installer.sh | sh
  ```

- [ ] **MSI installer (Windows)**: Download and test from GitHub Releases

## Post-Release

- [ ] Announce release (if applicable):
  - [ ] Reddit (r/rust, r/commandline)
  - [ ] Twitter/X
  - [ ] Hacker News (for major releases)
- [ ] Update README.md if installation instructions changed
- [ ] Close release tracking issue

## Troubleshooting

If something goes wrong, see the **Troubleshooting Releases** section in AGENTS.md.

Common issues:
- **GitHub Actions fails**: Check workflow logs for specific error
- **Scoop installation broken**: Verify manifest uses `.zip` file, not `.msi`
- **AUR build fails**: Double-check source URL and SHA256 hash
- **Homebrew formula outdated**: Check homebrew-tap repo for commit
