Option Compare Database

' ----------------------------------------------------------------------
'
' This subroutine will create a new table in the current database.
' The table will comes from an ODBC source (like a SQL Server instance).
'
' You can choose of importing locally the data or just attach the table.
' You can also choose to use your own username or choose an application login.
'
' Parameters : 
'    @bImport      : True will means that the table will be imported locally in 
'                        your MS Access DB. No link will be keept with the ODBC
'                        database.
'                    False will means that, in term of MS Access, an attached table
'                        will be created and, therefore, data won't be copied 
'                        locally so you'll always have the latest versions 
'                        ("attach" means online)
'    @sServerName  : Name of your ODBC servername (where the DB is stored)
'    @sDBName      : Name of your ODBC database name 
'    @sSourceTable : Name, in the ODBC database, of the table to link / import
'    @sLocalTable  : Name, in your current MS Access database, of the table to 
'                        create, locally
'    @bTrusted     : True means that the connection will be made with your current
'                        Windows credentials
'                    False means that the connection must be made with 
'                        specific credentials and not yours. An application login f.i.
'    @sUserID      : In case of bTrusted=False, user to use for the connection
'    @sPassword    : In case of bTrusted=False, password to use for the connection
'
' Example : 
'
'	'Link
'	Call Attach_Table(True, "servername", "dbName", "dbo.Data", "tblData", True)
'
'	'Import and user specific credentials
'	Call Attach_Table(False, "servername", "dbName", "dbo.Data", "tblData", False, _
'		"username", "password")
'
' ----------------------------------------------------------------------

Sub Attach_Table(ByVal bImport As Boolean, ByVal sServerName As String, _
    ByVal sDBName As String, ByVal sSourceTable As String, ByVal sLocalTable As String, _
    Optional ByVal bTrusted As Boolean = True, Optional ByVal sUserID As String, _
    Optional ByVal sPassword As String)
   
Dim sCredentials As String
Dim wType As Byte

    If (bImport) Then
        wType = 0 ' acImport
    Else
        wType = 2 ' acLink
    End If

    If bTrusted Then
        sCredentials = "Trusted_Connection=Yes;"
    Else
        sCredentials = "UID=" & sUserID & ";PWD=" & sPassword & ";"
    End If
   
    DoCmd.TransferDatabase wType, "ODBC Database", _
        "ODBC;Driver={SQL Server};Server=" & sServerName & ";" & _
        "Database=" & sDBName & ";" & sCredentials, _
        acTable, sLocalTable, sSourceTable
        
End Sub

Sub Test()

    ' Import the table
	'
    '    Using trusted connection
    'Call Attach_Table(True, "vwebudget.yourict.net", "eBudget", "dbo.Bistel", "Bistel", True)
	'
    '    Using application's login
    'Call Attach_Table(True, "vwebudget.yourict.net", "eBudget", "dbo.Bistel", "Bistel", _
	'	False, "userEBUDLaw_R", "v9lVd6QzVDX6DSDlyiuJ")
        
    ' Attach the table
	'
    '    Using trusted connection
    'Call Attach_Table(False, "vwebudget.yourict.net", "eBudget", "dbo.Bistel", "Bistel", _
	'    True)
	'
    '    Using application's login
    'Call Attach_Table(False, "vwebudget.yourict.net", "eBudget", "dbo.Bistel", _
	'    "Bistel", False, "userEBUDLaw_R", "v9lVd6QzVDX6DSDlyiuJ")
   
End Sub
