[![CI](https://github.com/hkwk/dtEEMCG/actions/workflows/ci.yml/badge.svg?branch=master)](https://github.com/hkwk/dtEEMCG/actions)

**CI:** GitHub Actions runs `cargo fmt`, `cargo clippy` and `cargo test` on push and PRs. See `.github/workflows/ci.yml`.

A Rust toolbox for Excel transformations.

This package builds two binaries:

- `dtEEMCG`: VOCs/NMHC sheet rename + cell edits
- `dtproton`: Ion chromatography data processing and formatting

## dtproton

The `dtproton` binary processes Excel files containing ion chromatography data and transforms them into a standardized output format.

### Input Format

Provisional Environment Monitoring Data

### Output Format

The output Excel file is generated with the CNEMC Air Monitoring Data Format.

### Data Processing Rules

1. **Time formatting**: Time values from the input are formatted to "YYYY-MM-DD HH:MM:SS" format, preserving the original date
2. **Data filtering**: Cells containing "(C)" or "(RM)" identifiers are set to empty
3. **Non-numeric values**: Cells containing non-numeric strings (such as "—", "N/A", etc.) are set to empty
4. **Column mapping**: Ion concentration data is mapped to the correct columns

### Configuration File

The `dtproton` binary uses a configuration file `proton_config.txt` to store the text content for Row 2 of the output file. This allows you to customize the instructions without modifying the code.

- If `proton_config.txt` exists in the working directory, the content will be read from this file
- If the file does not exist, a default text will be used
- The configuration file is excluded from version control (see `.gitignore`)

To create your own configuration file:

1. Copy `proton_config.example.txt` to `proton_config.txt`
2. Edit `proton_config.txt` with your desired content
3. The content will be used for Row 2 of the output file

### Usage

```bash
cargo run --bin dtproton -- <input.xlsx>
```

The output file will be saved as `processed_<input.xlsx>` in same directory.

## dtEEMCG

The `dtEEMCG` binary handles VOCs/NMHC sheet renaming and cell edits.

### Usage

```bash
cargo run --bin dtEEMCG -- <input.xlsx>
```

## Generate Sample Workbook

```bash
cargo run --example gen_sample
```
