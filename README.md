# PowerShell WiFi Profiles Exporter

This PowerShell script allows you to export WiFi profiles, including clear-text passwords, from your system asynchronously using RunspacePool. It enhances performance by leveraging parallel execution, making it ideal for systems with multiple WiFi profiles.

## Features

- **Asynchronous Execution**: Export WiFi profiles concurrently to speed up the process.
- **Thread Management**: Utilizes RunspacePool for efficient thread management.
- **Export Formats**: Generates XML and CSV files containing WiFi profile details.

## Usage

1. Clone this repository to your local machine:
   ```bash
   git clone https://github.com/RangO1972/PowerShell-WiFi-Profiles-Exporter.git
   ```
2. Open PowerShell and navigate to the directory containing the script (Export-WiFiProfiles.ps1).
3. Run the script:
   ```powershell
   .\Export-WiFiProfiles.ps1
   ```

## Requirements
* Windows PowerShell 4.0 or higher.
  
## Notes
This script does not support all edge cases and may require adjustments based on your system configuration.
Always review exported files (*.xml and *.csv) for accuracy and security considerations.
Contributing
Contributions are welcome! Feel free to fork the repository, make improvements, and submit pull requests.

## License
This project is licensed under the MIT License - see the LICENSE file for details.
