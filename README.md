# ExcelDataExtract

This is a .NET Core project to upload the Excel workbooks and extract the data from it and create dynamic tables in MS SQL Server and upload the data in to those tables


## Table of Contents

- [Getting Started](#getting-started)
  - [Prerequisites](#prerequisites)
  - [Installation](#installation)
- [Usage](#usage)

## Getting Started

1. Create a Test Database:
Begin by creating a test database in your Microsoft SQL Server.

2. Configure appsettings.json:
Open the appsettings.json file and make the following modifications within the ConnectionStrings section:
Update the "constr" value with your MS SQL Server server details, including the Data Source, Initial Catalog (Database Name), and Integrated Security settings.

3. Configure Excel Connection String:
Still in appsettings.json, modify the "ExcelConString" value under ConnectionStrings with your Microsoft Access Database Engine version and Excel version. Ensure the placeholder {0} in the connection string is appropriately replaced

### Prerequisites

Visual Studio,.NET SDK installed on your machine(.NET, .NET Core), Microsoft SQL Server,Micorosft AccessDatabaseEngine 12.0 or higher, Microsft Excel installed on your machine, A sample Excel data to upload extract data, NuGet Package Manager. 

### Installation
Ensure the presence of the following NuGet packages in your project. If any of these packages are missing, install them using the Package Manager Console or Visual Studio NuGet Package Manager.
** System.Data.SqlClient -> Install-Package System.Data.SqlClient
** System.Data.OleDb -> Install-Package System.Data.OleDb
** Serilog -> Install-Package Serilog
** Serilog.Sinks.File -> Install-Package Serilog.Sinks.File
** Serilog.Sinks.Console -> Install-Package Serilog.Sinks.Console
** Serilog.Extensions.Logging -> Install-Package Serilog.Extensions.Logging


## Usage

Follow these simple steps to use the Excel Export functionality in this application:

1. Build and Run:
Build and run the application on your ASP.NET Core environment.

2. Navigate to Excel Export:
In the application, go to the 'Excel Export' section in the navigation bar.

3. Choose File:
Click on 'Choose file' to select the Excel file you want to import.

4. Import Data:
After choosing the file, click on 'Import' to initiate the data extraction process.

5. Check Status Message:
Upon completion, a success message will be displayed in green if the data extraction and table creation were successful.

6. Troubleshooting:
If the message appears in red, there might be an issue. Check the log file and review the code for more details on any errors.

7. File Naming Convention:
Ensure that the uploaded file does not contain unwanted characters such as '@', '#', '$', '%', '^', '&', '*', '!', '-', etc. Spaces in the file name are acceptable.

"By following these steps, you can seamlessly import Excel data, create dynamic tables, and receive clear status messages for successful or unsuccessful operations. For additional insights, refer to the log file and review the application code."
