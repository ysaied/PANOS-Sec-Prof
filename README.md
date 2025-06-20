# Panorama Security Rules Extractor

A Python script to extract pre- and post- security rules from all device-groups in a Palo Alto Networks Panorama and compile them into a formatted Excel report.

## Features

- Fetches all device-groups configured on the Panorama.
- Retrieves both **pre-rulebase** and **post-rulebase** security policies for each device-group.
- Extracts key information:
  - Rule name
  - Action (allow, deny, etc.)
  - Security profile group (if configured)
- Exports the collected data into an Excel spreadsheet (`.xlsx`) with:
  - Bold headers
  - Auto-sized columns
  - Frozen header row and autofilter for easy navigation
- Implements robust error handling and comprehensive logging.

## Prerequisites

- **Python 3.6+**
- **pip3** package manager

## Required Python Libraries

Install the non-stdlib dependencies with:

```bash
pip3 install requests xmltodict openpyxl
```

## Installation

1. Clone or download this repository to your Panorama management host.
2. Ensure the shebang is at the top of the script for portability:
   ```bash
   #!/usr/bin/env python3
   ```
3. Make the main script executable:
   ```bash
   chmod +x main.py
   ```

## Configuration

1. **credentials.py**
   Create a file named `credentials.py` in the same directory with the following content:
   ```python
   # credentials.py
   api_key = "YOUR_PANORAMA_API_KEY"
   ```
   Replace `YOUR_PANORAMA_API_KEY` with the API key generated for your Panorama user.

2. **Script Variables**
   In `main.py`, update the `PANORAMA_IP` constant to point to your Panorama management IP or hostname:
   ```python
   PANORAMA_IP = "192.168.1.1"
   ```

## Usage

Run the script directly:

```bash
./main.py
```

Or explicitly via Python:

```bash
python3 main.py
```

The script will:

1. Retrieve all device-groups.
2. Fetch pre- and post-security rules for each group.
3. Compile the data and generate `Panorama_Security_Rules.xlsx` in the current directory.
4. Log progress and any errors to `panorama_rules.log`.

## Output

- **Panorama_Security_Rules.xlsx**: Excel file with columns:
  - Device Group
  - Pre/Post
  - Rule Name
  - Action
  - Security Profile

- **panorama_rules.log**: Log file capturing the script’s execution details and any errors.

## Logging

The script uses Python’s `logging` module to log:

- INFO messages for high-level progress (e.g., number of device-groups retrieved).
- DEBUG messages for API request details.
- ERROR messages if API calls or file operations fail.

Logs are output to both the console and `panorama_rules.log`.

## Troubleshooting

- **Missing rules**: If a device-group has no pre- or post-rules, it will be skipped with an INFO log.
- **API Errors**: Check `panorama_rules.log` for any authentication failures or unexpected API responses.
- **Excel Save Issues**: Ensure you have write permissions in the target directory.

## Extending the Script

- Add additional columns (e.g., source/destination zones) by modifying `extract_profile_name()` and the data row assembly in `main.py`.
- Filter rules by action or profile by adding logic in the rule-processing loop.
- Schedule this script via cron or task scheduler for regular reporting.

## License

This project is released under the MIT License. See the [LICENSE](LICENSE) file for details.
