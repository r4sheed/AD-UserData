# Export-ADUserData PowerShell Script

This PowerShell script exports Active Directory user data to a specified file format (CSV or TXT). It supports filtering users, customizing output, and exporting data from a specific organizational unit (OU).

---

## Overview

The `Export-ADUserData` function retrieves user information from Active Directory, processes the data, and exports it to a file in either CSV or TXT format. It allows you to exclude specific users using regex patterns and customize the search base for the directory.

---

## How to Use

1. **Prerequisites**:
   - Ensure the Active Directory module (`ActiveDirectory`) is installed and imported.
   - You must have the necessary permissions to query Active Directory and write to the specified output directory.

2. **Run the Script**:
   - Copy the script into a `.ps1` file or import it into your PowerShell session.
   - Execute the `Export-ADUserData` function with the desired parameters.

---

## Parameters

| Parameter     | Description                                                                 | Required | Default Value                     |
|---------------|-----------------------------------------------------------------------------|----------|-----------------------------------|
| `OutputPath`  | The file path for the exported data. Supported extensions: `.csv`, `.txt`.  | Yes      | -                                 |
| `Format`      | The file format of the output. Accepted values: `CSV`, `TXT`.               | No       | `CSV`                            |
| `SkipUsers`   | A list of usernames (supports regex patterns) to exclude from the export.   | No       | `@()` (empty array)              |
| `SearchBase`  | The distinguished name of the directory search base.                        | No       | `"OU=Company,DC=company,DC=com"` |

---

## Examples

### Export User Data to a CSV File
```powershell
Export-ADUserData -OutputPath "C:\Reports\UserData.csv"
```

### Export User Data to a TXT File While Excluding Specific Users
```powershell
Export-ADUserData -OutputPath "C:\Reports\UserData.txt" -Format TXT -SkipUsers @('jdoe.*', 'admin*')
```

### Export User Data from a Specific Organizational Unit
```powershell
Export-ADUserData -SearchBase "OU=Users,OU=HQ,DC=company,DC=com" -OutputPath "C:\Reports\HQUserData.csv"
```

---

## Notes

- **Active Directory Module**: This script requires the `ActiveDirectory` module to be installed and imported. You can install it using:
  ```powershell
  Install-WindowsFeature -Name RSAT-AD-PowerShell
  ```

- **File Permissions**: Ensure you have the necessary permissions to create or overwrite files in the specified output directory.

- **Error Handling**: The script includes error handling to catch and report issues during execution.

---

## Output

The script writes the exported user data to the specified output file. The output includes the following fields:

- `sAMAccountName`
- `Name`
- `Mobile`
- `TelephoneNumber`
- `Mail`
- `Description`
- `Skype`
- `dxMidOU`
- `IsVezeto`
- `IsHidden`

---

## License

This project is licensed under the MIT License. See the [LICENSE](LICENSE) file for details.

---

## Contributing

Contributions are welcome! Please open an issue or submit a pull request for any improvements or bug fixes.

---

## Support

For questions or issues, please open an issue in the GitHub repository.