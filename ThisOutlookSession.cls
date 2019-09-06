VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisOutlookSession"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private WithEvents sentItems As Outlook.Items
Attribute sentItems.VB_VarHelpID = -1
Private dbPath As String

Const PR_SMTP_ADDRESS As String = "http://schemas.microsoft.com/mapi/proptag/0x39FE001E"

'defined what folder we need to monitor
Sub defFolder()
  
  Dim outlookApp As Outlook.Application
  Dim objectNS As Outlook.NameSpace

  Set outlookApp = Outlook.Application
  Set objectNS = outlookApp.GetNamespace("MAPI")
  Set sentItems = objectNS.GetDefaultFolder(olFolderSentMail).Items

End Sub

'run the function when someone a mail has been send out
Private Sub Application_ItemSend(ByVal Item As Object, Cancel As Boolean)

  defFolder

End Sub

'run the function on outlook startup
Private Sub Application_Startup()

  defFolder
 
End Sub


'This code will monitor all mails added in sent item
Private Sub sentItems_ItemAdd(ByVal Item As Object)

'check if item is a mail
If Item.Class = olMail Then

    Dim recips As Outlook.Recipients, recip As Outlook.Recipient
    Dim recipPA As Outlook.PropertyAccessor, sendPA As Outlook.PropertyAccessor
    Dim strIDs() As String, intX As Integer

    'get sender name and email address
    Set sendPA = Item.Sender.PropertyAccessor
    
    'get reciever name and email address
    recName = ""
    recMail = ""
    Set recips = Item.Recipients
    For Each recip In recips
        Set recipPA = recip.PropertyAccessor
        recName = recip.Name & "," & recName
        recMail = recipPA.GetProperty(PR_SMTP_ADDRESS) & "," & recMail
    Next

    eType = checkEntry(Item.ConversationID)

    If eType = "" Then

        dbPath = dbFolder & Format(Date, "YYYY") & "\" & Format(Date, "MMMM") & ".accdb"

        query = "INSERT INTO Email_Details VALUES (" & _
                    "'" & Item.ConversationID & "'," & _
                    "'" & Item.Sender.Name & "'," & _
                    "'" & sendPA.GetProperty(PR_SMTP_ADDRESS) & "'," & _
                    "'" & Replace(Left(recName, Len(recName) - 1), "'", "") & "'," & _
                    "'" & Left(recMail, Len(recMail) - 1) & "'," & _
                    "'" & Item.Subject & "'," & _
                    "'" & Item.SentOn & "'," & _
                    "'" & Item.SentOn & "'," & _
                    "'" & Item.SentOn & "'," & _
                    "'Open'," & _
                    "'Outgoing');"
        mySQLRun query, dbPath

        query = "INSERT INTO Email_Content VALUES (" & _
                    "'" & Item.ConversationID & "'," & _
                    "'" & Item.body & "');"

        mySQLRun query, dbPath
    Else

        If eType = "Incoming" Then
            query = "UPDATE Email_Details " & _
                    "SET Last_Touch = '" & Item.SentOn & "', " & _
                        "Replied_Date = '" & Item.SentOn & "', " & _
                        "Status = 'Close' " & _
                    "WHERE Conversation_ID = '" & Item.ConversationID & "';"
            mySQLRun query, dbPath

        Else
            query = "UPDATE Email_Details " & _
                    "SET Last_Touch = '" & Item.SentOn & "', " & _
                        "Status = 'Open' " & _
                    "WHERE Conversation_ID = '" & Item.ConversationID & "';"
            mySQLRun query, dbPath
        End If

        query = "UPDATE Email_Content " & _
                "SET Email_Body = '" & Item.body & "' " & _
                "WHERE Conversation_ID = '" & Item.ConversationID & "';"
        mySQLRun query, dbPath

    End If

    query = "INSERT INTO Email_Log VALUES (" & _
                "'" & Item.ConversationID & "', " & _
                "'" & Item.SentOn & "', " & _
                "'Outgoing');"
    mySQLRun query, dbPath

End If

End Sub

'This will monitor all incoming mails

Private Sub Application_NewMailEx(ByVal EntryIDCollection As String)

Dim eArray As Variant
Dim objNS As Object, objEmail As Object 'Outlook.MailItem
Dim recips As Outlook.Recipients, recip As Outlook.Recipient
Dim recipPA As Outlook.PropertyAccessor, sendPA As Outlook.PropertyAccessor
Dim strIDs() As String, intX As Integer
Dim recName As String, recMail As String, query As String, eType As String, eAddress As String

eArray = Array("notification@fbworkmail.com", "XPOCFWC@xpo.com", "itcommunications@xpo.com", "informationsecurity@xpo.com", "XPOITFeedback@xpo.com", "XPOITServiceDesk@xpo.com")

strIDs = Split(EntryIDCollection, ",")

For intX = 0 To UBound(strIDs)

    Set objNS = Application.GetNamespace("MAPI")
    Set objEmail = objNS.GetItemFromID(strIDs(intX))

    If objEmail.Class = olMail Then

        'get sender name and email address
        If objEmail.SenderEmailType = "EX" Then
            Set sendPA = objEmail.Sender.PropertyAccessor
            eAddress = sendPA.GetProperty(PR_SMTP_ADDRESS)
        Else
            eAddress = objEmail.SenderEmailAddress
        End If

        If UBound(Filter(eArray, eAddress)) = -1 Then
        
            'get reciever name and email address
            Set recips = objEmail.Recipients
            recName = ""
            recMail = ""
            For Each recip In recips
                Set recipPA = recip.PropertyAccessor
                recName = recip.Name & "," & recName
                recMail = recipPA.GetProperty(PR_SMTP_ADDRESS) & "," & recMail
            Next
    
            eType = checkEntry(objEmail.ConversationID)
            
            If eType = "" Then
            
                dbPath = dbFolder & Format(Date, "YYYY") & "\" & Format(Date, "MMMM") & ".accdb"
                
                query = "INSERT INTO Email_Details VALUES (" & _
                            "'" & objEmail.ConversationID & "'," & _
                            "'" & objEmail.Sender.Name & "'," & _
                            "'" & eAddress & "'," & _
                            "'" & Replace(Left(recName, Len(recName) - 1), "'", "") & "'," & _
                            "'" & Left(recMail, Len(recMail) - 1) & "'," & _
                            "'" & objEmail.Subject & "'," & _
                            "'" & objEmail.ReceivedTime & "'," & _
                            "'" & objEmail.ReceivedTime & "'," & _
                            "'" & objEmail.ReceivedTime & "'," & _
                            "'Open'," & _
                            "'Incoming');"
                mySQLRun query, dbPath
    
                query = "INSERT INTO Email_Content VALUES (" & _
                            "'" & objEmail.ConversationID & "'," & _
                            "'" & objEmail.body & "');"
                mySQLRun query, dbPath
                
            Else
    
                If eType = "Incoming" Then
                    query = "UPDATE Email_Details " & _
                            "SET Last_Touch = '" & objEmail.ReceivedTime & "', " & _
                                "Status = 'Open' " & _
                            "WHERE Conversation_ID = '" & objEmail.ConversationID & "';"
                    mySQLRun query, dbPath
                Else
                    query = "UPDATE Email_Details " & _
                            "SET Last_Touch = '" & objEmail.ReceivedTime & "', " & _
                                "Replied_Date = '" & objEmail.ReceivedTime & "', " & _
                                "Status = 'Close' " & _
                            "WHERE Conversation_ID = '" & objEmail.ConversationID & "';"
                    mySQLRun query, dbPath
                End If
    
                query = "UPDATE Email_Content " & _
                        "SET Email_Body = '" & objEmail.body & "' " & _
                        "WHERE Conversation_ID = '" & objEmail.ConversationID & "';"
                mySQLRun query, dbPath
    
            End If
    
            query = "INSERT INTO Email_Log VALUES (" & _
                        "'" & objEmail.ConversationID & "', " & _
                        "'" & objEmail.ReceivedTime & "', " & _
                        "'Incoming');"
            mySQLRun query, dbPath
            
        End If
        
    End If

Next

End Sub


Sub createDB(path)
    
    Dim objCatalog As Object, sqlCode As String
    
    Set objCatalog = CreateObject("ADOX.Catalog")
    objCatalog.Create "PROVIDER=Microsoft.ACE.OLEDB.12.0;DATA SOURCE =" & path & ";"
    Set objCatalog = Nothing
    
    sqlCode = "CREATE TABLE Email_Details (" & _
                            "Conversation_ID VARCHAR(50), " & _
                            "Sender_Name VARCHAR(100), " & _
                            "Sender_Email_Address VARCHAR(100), " & _
                            "Recipient_Name TEXT, " & _
                            "Recipient_Email_Address TEXT, " & _
                            "Email_Subj VARCHAR(255), " & _
                            "Received_Date DATETIME, " & _
                            "Replied_Date DATETIME, " & _
                            "Last_Touch DATETIME, " & _
                            "Status VARCHAR(10), " & _
                            "Email_Type VARCHAR(10));"

    mySQLRun sqlCode, path
    
    sqlCode = "CREATE TABLE Email_Content (" & _
                            "Conversation_ID VARCHAR(50), " & _
                            "Email_Body TEXT);"
                            
    mySQLRun sqlCode, path

    sqlCode = "CREATE TABLE Email_Log (" & _
                            "Conversation_ID VARCHAR(50), " & _
                            "Log_Date DATETIME, " & _
                            "Email_Body VARCHAR(10));"
                            
    mySQLRun sqlCode, path
    
End Sub


Function checkEntry(convID)

Dim thisYear As String, thisMonth As String
Dim qry As String, eType As String, folderPath As String

thisYear = Format(Date, "YYYY")
thisMonth = Format(Date, "MMMM")

folderPath = dbFolder & thisYear
dbPath = dbFolder & thisYear & "\" & thisMonth & ".accdb"

If Dir(folderPath, vbDirectory) = Empty Then
    MkDir folderPath
End If

If Len(Dir(dbPath)) = 0 Then
    createDB dbPath
End If

qry = "SELECT Email_Type FROM Email_Details " & _
      "WHERE Conversation_ID = '" & convID & "';"
      
eType = mySQLReq(qry, dbPath)

If eType = "" Then
    backMonth = Format(DateAdd("D", -10, Date), "MMMM")
    If Not thisMonth = backMonth Then
        Dim prevYear As String, prevMonth As String
        
        prevMonth = Format(DateAdd("M", -1, Date), "MMMM")
        prevYear = Format(DateAdd("Y", -1, Date), "YYYY")
        
        If prevMonth = "December" Then
            dbPath = dbFolder & prevYear & "\" & prevMonth & ".accdb"
        Else
            dbPath = dbFolder & thisYear & "\" & prevMonth & ".accdb"
        End If
        
        If Len(Dir(dbPath)) = 0 Then
            checkEntry = eType
        Else
            checkEntry = mySQLReq(qry, dbPath)
        End If
    Else
        checkEntry = eType
    End If
Else
    checkEntry = eType
End If

End Function







