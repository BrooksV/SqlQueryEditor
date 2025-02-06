# SqlQueryEditor

A PowerShell WPF (Windows Presentation Framework) application to create, store, and Execute SQL Queries from SQL Database. tables.

The `SqlQueryEditor` is currently under development and many changes are expected over the next few months.

It's being designed as a Helper App that demonstrates how to build PowerShell WPF GUI interfaces to create, store, update, and delete records in database tables.

It could be used stand-alone or integrated into existing PS WPF Apps that use SQL Table data.

The challange is to avoid resorting to using a Model-View approach with ObservableCollection. This approach requires lots of hard coding of database schemas as classes and methods to read, update, and save data to the database tables.

The approach `SqlQueryEditor` will attempt is to bind the WPF dta elements to DataViews. Then leverage the dynamic nature of DataAdapters and SqlBuilder to get and create System.Data.SqlClient.DeleteCommand]DeleteCommand, System.Data.SqlClient.InsertCommand]InsertCommand, System.Data.SqlClient.UpdateCommand]UpdateCommand, and [System.Data.SqlClient.SqlCommand}SelectCommand. With changes being made to the WPF data elements, the Save button action will apply the DataGrid changes, the bound DataView is copied and filtered to post the changes back to the database.

# Requirements

## SqlQueryClass Module

SqlQueryClass Module is required to mantain and manage SQL Queries and results that are used in binding the data source to WPF controls.

### Installation of the SqlQueryClass Module

Find-Module -Name SqlQueryClass | Install-Module -Scope CurrentUser -AcceptLicense



