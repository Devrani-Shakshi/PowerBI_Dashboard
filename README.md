# Power BI Integration Dashboard

A professional ASP.NET Core web application designed to manage, sync, and embed Power BI reports and workspaces programmatically. This project leverages the Power BI REST APIs and Service Principal authentication to provide a seamless administrative and viewing experience.

##  Features

- **Workspace Management**: Programmatically create, rename, and delete Power BI workspaces.
- **Bidirectional Sync**: Keep your local database in sync with Power BI workspaces and reports.
- **Report Management**: 
  - Upload `.pbix` (Power BI) and `.rdl` (Paginated) files directly to the service.
  - Automatic detection of report types.
- **Embedded Analytics**: Securely embed reports into your application using Service Principal-based token generation.
- **Schema Discovery**: Automatically discover tables, columns, and parameters from Power BI datasets for dynamic filtering.
- **Export Capabilities**: Export reports to PDF format programmatically.
- **Modern UI**: A responsive dashboard built for administrative efficiency.

##  Technology Stack

- **Backend**: ASP.NET Core 8.0, C#
- **ORM**: Entity Framework Core
- **Database**: SQL Server (LocalDB / Express)
- **API Integration**: Microsoft Power BI API SDK
- **Authentication**: Azure AD (MSAL.NET) with Service Principal

##  Setup & Configuration

### Prerequisites
- .NET 8.0 SDK
- SQL Server (Express or LocalDB)
- Azure AD App Registration (Service Principal)
- Power BI Pro or Fabric/Premium Capacity (Required for some features like Export and RDL)

### Configuration
1.  Locate `PowerBI/appsettings.json.template`.
2.  Create a copy named `appsettings.json`.
3.  Fill in your Azure and Power BI credentials:

```json
{
  "PowerBI": {
    "TenantId": "your-tenant-id",
    "ClientId": "your-client-id",
    "ClientSecret": "your-client-secret",
    "AdminEmail": "your-admin-email",
    "CapacityId": "your-premium-capacity-id"
  },
  "ConnectionStrings": {
    "DefaultConnection": "Server=...;Database=PowerBIAppDB;..."
  }
}
```

### Installation
1.  Clone the repository:
    ```bash
    git clone https://github.com/Devrani-Shakshi/PowerBI_Dashboard.git
    ```
2.  Restore dependencies:
    ```bash
    dotnet restore
    ```
3.  Update the database:
    ```bash
    dotnet ef database update
    ```
4.  Run the application:
    ```bash
    dotnet run
    ```

