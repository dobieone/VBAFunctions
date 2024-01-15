# VBA Functions

A collection of VBA Helpers and Database Wrappers for Excel that I have used to automate office applications over the past 20 years.

## VBA References

The following references are required for the Helper Functions

- Microsoft Scripting Runtime

The following references are required for the Connection Wrappers

- Microsoft ActiveX Data Objects x.x Library
- Microsoft Scripting Runtime

The helper functions are also required for the Connection Wrappers

## VBA Helper Functions

- **Inc** - Increase a variable by an amount
- **Dec** - Decrease a variable by an amount
- **MakeID** - Makes a unique[ish] numeric ID based on the current date/time
- **IDToDate** - Converts a ID back to a datetime
- **LoadTextFile** - Loads a text file into a variable
- **BrowseFolder** - Shows the folder browser and returns the selected folder
- **BrowseFile** - Shows the file browser and returns the selected file
- **LastRow** - Return the last row used
- **LastCol** - Return the last column used
- **CreateNamedRange** - Creates a named range on a sheet
- **NamedRangeExists** - Checks to see if a named range exists on a sheet
- **UpdateChartAxis** - Updates a charts min/max axis values

## Data Connection Wrapper

A wrapper that allows you to connect to different data sources using the same logic. The wrapper does not support Update/Execution querys.

Sources connections available:

- **Excel** - xlsx, xlsm, xlsb, xls (untested)
- **MS Access** - accdb, accdr, mdb (untested)
- **Text Files** - txt, csv (schema.ini may be required for all text files)

### Example Usage
```vb
Public Sub LoadDataFromAccess()

    Dim db As New DataConnect
    
    With db
        ' Reset the class to defaults
        .Reset
        
        ' Set the Database name and path
        .Databasename = "TestDatabase.accdb"
        .DatabasePath = ActiveWorkbook.Path & "\"
        
        ' Load query from a file or set it manually
        .QueryFileName = "Query.sql"
        .QueryPath = ActiveWorkbook.Path & "\"
        ' .SQL = "SELECT * FROM TestTable where MyID = {ID}"
        
        ' BuildQuery will replace anything in {} with "keys" added here
        .AddKey "ID", "MyID"

        ' Set the output tab        
        .ResultsTab = "Sheet1"
        
        .BuildQuery
        
        ' Connect to source and run query
        .Connect
        .Run
        .Disconnect
    
    End With
    
    ' Clean up
    Set db = Nothing

End Sub
```

## Server Connection Wrapper

A wrapper that allows you to connect to SQL Server and load data to excel sheets. The wrapper does not support Update/Execution querys.

### Example Usage
```vb
Public Sub LoadDataFromAccess()

    Dim db As New ServerConnect
    
    With db
        ' Reset the class to defaults
        .Reset
        
        ' Set the Server name and Database name
        .ServerName = "MySQLServerName"
        .Databasename = "MyDBName"
        
        ' Load query from a file or set it manually
        .QueryFileName = "Query.sql"
        .QueryPath = ActiveWorkbook.Path & "\"
        ' .SQL = "SELECT * FROM TestTable where MyID = {ID}"
        
        ' BuildQuery will replace anything in {} with "keys" added here
        .AddKey "ID", "MyID"

        ' Set the output tab        
        .ResultsTab = "Sheet1"
        
        .BuildQuery
        
        ' Connect to source and run query
        .Connect
        .Run
        .Disconnect
    
    End With
    
    ' Clean up
    Set db = Nothing

End Sub
```