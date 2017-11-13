# Attach_Table

MS Access VBA code to make easy to create, dynamically, a table in the current database.

The link or import is made by using an ODBC connection.

See [module.bas](module.bas) for full source code

## Description

This subroutine will create a new table in the current database.
The table will comes from an ODBC source (like a SQL Server instance).

You can choose of importing locally the data or just attach the table.
You can also choose to use your own username or choose an application login.

## Parameters for Attach_Table() 

* @bImport : 
	**True** will means that the table will be imported locally in your MS Access DB. No link will be keept with the ODBC database.
	**False** will means that, in term of MS Access, an attached table will be created and, therefore, data won't be copied locally so you'll always have the latest versions ("attach" means online)
* @sServerName : Name of your ODBC servername (where the DB is stored)
* @sDBName : Name of your ODBC database name 
* @sSourceTable : Name, in the ODBC database, of the table to link / import
* @sLocalTable : Name, in your current MS Access database, of the table to create, locally
* @bTrusted : 
	**True** means that the connection will be made with your current Windows credentials
	**False** means that the connection must be made with specific credentials and not yours. An application login f.i.
* @sUserID : In case of bTrusted=False, user to use for the connection
* @sPassword : In case of bTrusted=False, password to use for the connection

## Sample code : 

```VB
	'Link
	Call Attach_Table(True, "servername", "dbName", "dbo.Data", "tblData", True)

	'Import and user specific credentials
	Call Attach_Table(False, "servername", "dbName", "dbo.Data", "tblData2", False, _
		"username", "password")
```
