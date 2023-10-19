# MEM_CreateGroups

Script designed to automate the creation of Azure AD (Microsoft Entra) Dynamic Groups. Uses a CSV file to define group attributes and can operate across multiple tenants. 
I use script in Microsoft Intune projects to to automate group creation of base or initial groups.

### Features

- **CSV-Driven Group Creation**: Utilizes a CSV file to define group attributes such as name, description, membership rules, and more.
  
- **Multi-Tenant Support**: The script allows for the specification of a Tenant ID, enabling operations across different Azure AD tenants.

- **Module Validation**: Validates and installs the required AzureAD or Az PowerShell modules.

- **Error Handling**: Robust error checking and validation mechanisms are in place.

### Prerequisites

- Administrative privileges on the machine where the script is run.
- AzureAD or Az PowerShell module.

### Usage

#### With Tenant ID
```powershell

.\MEM_CreateGroups.ps1 -GroupsFile .\YourGroupsFile.csv -TenantId "your-tenant-id"

```

#### Without Tenant ID
```powershell

.\MEM_CreateGroups.ps1 -GroupsFile .\YourGroupsFile.csv

```

Certainly! Here's the updated section on the CSV file format, based on the details from the original script:

### CSV File Format

The CSV file should contain the following headers:

- `GroupType`: The type of the group (e.g., Security, Office365).
- `GroupDisplayName`: The display name for the group.
- `GroupDescription`: A description for the group.
- `GroupMembershipType`: The membership type. Acceptable values are "AA" for assigned groups, "DD" for Dynamic Device groups, and "DU" for Dynamic User groups.
- `GroupMembershipRule`: The membership rule in OData format.
- `GroupOwner`: The owner of the group (optional).

Here's a sample CSV file content:

```csv
GroupType,GroupDisplayName,GroupDescription,GroupMembershipType,GroupMembershipRule,GroupOwner
Security,HR Team,Human Resources Team,DU,(user.department -eq "HR"),john.doe@domain.com
```



### Output

- **Azure AD Groups**: Dynamic groups are created in the specified or default Azure AD tenant.

### Contributing

Contributions to the MEM_CreateGroups repository are welcomed. If you have a feature request, bug report, or improvement, please open an issue or submit a pull request.



### Disclaimer

This script is provided as-is with no warranties or guarantees. Always test in a controlled environment before production deployment.


