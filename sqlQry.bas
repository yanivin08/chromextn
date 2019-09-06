Attribute VB_Name = "sqlQry"
Const fileName = "\\INPCFFS0448.amer.corp.xpo.com\User\axsalaysaybayani@CONWAY\Desktop\EMS\Titan_EMS (Copy).accdb"
Public Const dbFolder = "\\cdcts002\sdrive\MAJOR ACCOUNT COLLECTIONS\Project 180\EMS Folder\"

Public Sub mySQLRun(sqlCode, path)
    'function to connect in sql database. Limited only in adding, update, deleting and creating data.
    
    Dim objConn As Object
    Set objConn = CreateObject("ADODB.Connection")
    objStr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & path & ";Persist Security Info=False;"
    objConn.Open objStr
    On Error GoTo errHandler
    objConn.Execute sqlCode
    objConn.Close
    Set objConn = Nothing
    
    Exit Sub
 
errHandler:
    MsgBox Err.Description

    
End Sub


Public Function mySQLReq(query, path)

    'This would check if the mail item has been added to the database
    Dim objConn As Object
    Dim objRecs As Object
    Dim sConn As String
    
    sConn = "Provider=Microsoft.ACE.OLEDB.12.0;" & _
                               "Data Source=" & path & ";" & _
                               "Jet OLEDB:Engine Type=5;" & _
                               "Persist Security Info=False;"

    Set objConn = CreateObject("ADODB.Connection")
    Set objRecs = CreateObject("ADODB.Recordset")
    
    objConn.ConnectionString = sConn
    objConn.Open
    
    With objRecs
        .ActiveConnection = objConn
        .Source = query
        .Locktype = 1
        .CursorType = 0
        .Open
    End With
    
    If objRecs.EOF Then
        mySQLReq = ""
    Else
        mySQLReq = objRecs.Fields.Item(0).Value
    End If
        
    objConn.Close
    Set objConn = Nothing
    Exit Function

errHandler:
    MsgBox Err.Description

End Function

