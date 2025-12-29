[![CI](https://github.com/hkwk/dtEEMCG/actions/workflows/ci.yml/badge.svg?branch=master)](https://github.com/hkwk/dtEEMCG/actions)

**CI:** GitHub Actions runs `cargo fmt`, `cargo clippy` and `cargo test` on push and PRs. See `.github/workflows/ci.yml`.

A Rust toolbox for Excel transformations.

This package builds two binaries:

- `dtEEMCG`: VOCs/NMHC sheet rename + cell edits
- `dtproton`: Ion chromatography cleanup

Usage:

- `cargo run --bin dtEEMCG -- <input.xlsx>`
- `cargo run --bin dtproton -- <input.xlsx>`

Generate a sample workbook:

- `cargo run --example gen_sample`
