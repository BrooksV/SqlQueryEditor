<div align="center" width="100%">
    <h1>SqlQueryEditor</h1>
    <p>Module that create an instance of a PowerShell class which is used to execute SQL Queries and manages output as DataTable, DataAdapter, DataSet, SqlReader, or NonQuery result object.</p><p>
    <a target="_blank" href="https://github.com/BrooksV"><img src="https://img.shields.io/badge/maintainer-BrooksV-orange" /></a>
    <a target="_blank" href="https://github.com/BrooksV/SqlQueryEditor/graphs/contributors/"><img src="https://img.shields.io/github/contributors/BrooksV/SqlQueryEditor.svg" /></a><br>
    <a target="_blank" href="https://github.com/BrooksV/SqlQueryEditor/commits/"><img src="https://img.shields.io/github/last-commit/BrooksV/SqlQueryEditor.svg" /></a>
    <a target="_blank" href="https://github.com/BrooksV/SqlQueryEditor/issues/"><img src="https://img.shields.io/github/issues/BrooksV/SqlQueryEditor.svg" /></a>
    <a target="_blank" href="https://github.com/BrooksV/SqlQueryEditor/issues?q=is%3Aissue+is%3Aclosed"><img src="https://img.shields.io/github/issues-closed/BrooksV/SqlQueryEditor.svg" /></a><br>
</div>

# SqlQueryEditor

A PowerShell WPF (Windows Presentation Framework) application to create, store, and Execute SQL Queries from SQL Database. tables.

## Description

The `SqlQueryEditor` is currently under development and many changes are expected over the next few months.

It's being designed as a Helper App that demonstrates how to build PowerShell WPF GUI interfaces to create, store, update, and delete records in database tables.

It could be used stand-alone or integrated into existing PS WPF Apps that use SQL Table data.

The challange is to avoid resorting to using a Model-View approach with ObservableCollection. This approach requires lots of hard coding of database schemas as classes and methods to read, update, and save data to the database tables.

The approach `SqlQueryEditor` will attempt is to bind the WPF dta elements to DataViews. Then leverage the dynamic nature of DataAdapters and SqlBuilder to get and create System.Data.SqlClient.DeleteCommand]DeleteCommand, System.Data.SqlClient.InsertCommand]InsertCommand, System.Data.SqlClient.UpdateCommand]UpdateCommand, and [System.Data.SqlClient.SqlCommand}SelectCommand. With changes being made to the WPF data elements, the Save button action will apply the DataGrid changes, the bound DataView is copied and filtered to post the changes back to the database.

## Module Install

SqlQueryEditor is in early development phase. Please read through [ChangeLog](/CHANGELOG.md) for all updates.

Stable releases can be installed from the PowerShell Gallery:

```PowerShell
Install-Module -Name SqlQueryEditor -Verbose
```

To load a local build of the module, use `Import-Module` as follows:

```PowerShell
Import-Module -Name ".\dist\SqlQueryEditor\SqlQueryEditor.psd1" -Force -verbose
```

## Requirements

- Tested with PowerShell 5.1 and 7.5x
- No known dependencies for usage
- Module build process uses [Manjunath Beli's](https://github.com/belibug) [ModuleTools](https://github.com/belibug) module.

## ToDo

- [ ] Seek peer review and comments
- [ ] Integrate feedback
- [ ] Improve Documentation

## LONG DESCRIPTION

## SqlQueryEditor Module

SqlQueryEditor Module is required to mantain and manage SQL Queries and results that are used in binding the data source to WPF controls.

### Troubleshooting

## Folder Structure and Build Management

The folder structure of the SqlQueryEditor module is based on best practices for PowerShell module development and was initially created using [Manjunath Beli's](https://github.com/belibug) [ModuleTools](https://github.com/belibug) module. Check out his [Blog article](https://blog.belibug.com/post/ps-modulebuild) that explains the core concepts of ModuleTools.

The the following ModuleTools CmdLets used in the build and maintenance process. They need to be executed from project root:

- Get-MTProjectInfo -- returns hashatble of project configuration which can be used in pester tests or for general troubleshooting
- Update-MTModuleVersion -- Increments SqlQueryEditor module version by modifying the values in `project.json` or you can manually edit the json file.
- Invoke-MTBuild -- Run `Invoke-MTBuild -Verbose` to build the module. The output will be saved in the `dist` folder, ready for distribution.
- Invoke-MTTest -- Executes pester configuration (*.text.ps1) files in the `tests` folder

- To skip a test, add `-skip` in describe block of the Pester *.test.ps1 file to skip.

### Folder and Files
 
```powershell

```

All files and folders in the `src` folder, will be published Module.

All other folder and files in the `.\SqlQueryEditor` folder will resides in the [GitHub SqlQueryEditor Repository](https://github.com/BrooksV/SqlQueryEditor) except those excluded by inclusion in the `.\SqlQueryEditor\.gitignore` file.

### Project JSON File

The `project.json` file contains all the important details about your module, is used during the module build process, and helps to generate the SqlQueryEditor.psd1 manifest.

### Root Level and Other Files

- .gitignore -- List of file, folder, and wildcard specifications to ignore when publishing to GitHub repository
- GitHub_Action_Docs.md -- How to add GitHub Action WorkFlows to automate CI/CD (Continuous Integration/Continuous Deployment)
- LICENSE -- MIT License notice and copyright
- project.json -- ModuleTools project configuration file used to build the `SqlQueryEditor` module
- README.md -- Documentation (this) file for the `SqlQueryEditor` module
- .vscode\settings.json -- VS Code settings used during `SqlQueryEditor` module development

### archive Folder

`.\SqlQueryEditor\archive` is not used in this project. Its a temporary place / BitBucket to hold code snippets and files during development and is not part of the build.

### Dist (build output) Folder

Generated module is stored in `dist\SqlQueryEditor` folder, you can easily import it or publish it to PowerShell Gallery or repository.

### Src Folder

- All functions in the `public` folder are exported during the module build.
- All functions in the `private` folder are accessible internally within the module but are not exposed outside the module.
- All files and folder contained in the `resources` folder will be `dist\SqlQueryEditor` folder.

### Tests Folder

If you want to run any `pester` tests, keep them in `tests` folder and named *.test.ps1.

Run `Invoke-MTTest` to execute the tests.

## How the `SqlQueryEditor` Module Works

SqlQueryEditor.ps1

## Contributing

Contributions are welcome! Please fork the repository and submit a pull request with your changes. Ensure that your code adheres to the existing style and includes appropriate tests.

## License

This project is licensed under the MIT License. See the LICENSE file for details.

[BadgeIOCount]: https://img.shields.io/powershellgallery/dt/SqlQueryEditor?label=SqlQueryEditor%40PowerShell%20Gallery
[PSGalleryLink]: https://www.powershellgallery.com/packages/SqlQueryEditor/
[WorkFlowStatus]: https://img.shields.io/github/actions/workflow/status/BrooksV/SqlQueryEditor/Tests.yml

## SEE ALSO

New-SqlQueryDataSet
Get-Help

## KEYWORDS

SQL, Database, Query, SqlQueryDataSet
