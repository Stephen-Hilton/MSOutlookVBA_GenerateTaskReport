Attribute VB_Name = "mod_GenerateTaskReport"
Option Explicit

Public Sub GenerateEmail_TaskReport()
    Dim emailBody As String
    Dim dtWeekStart As Date
    Dim objMsg As MailItem
    Set objMsg = Application.CreateItem(olMailItem)
    
    'What week are we reporting on?
    If Weekday(Now()) <= 4 Then 'if on-or-before Wednesday, then report LAST week
        dtWeekStart = Int(Now() - Weekday(Now()) - 6)
    Else 'if Thursday or after, then THIS week
        dtWeekStart = Int(Now() - Weekday(Now()) + 1)
    End If
    
    ''''''''''''''''''''''
    'Change stuff here:
    ''''''''''''''''''''''
    With objMsg
        .To = "Nathan.Wenig@Teradata.com; Bryan.Owen@Teradata.com; Sharif.Islam@Teradata.com;"
        .CC = "Chirag.Belosay@Teradata.com; Farha.Adeni@Teradata.com; Stephen.Hilton@Teradata.com;"
        .Subject = "SA Weekly Update: Stephen Hilton"
        .Categories = "Update"
        .BodyFormat = olFormatHTML
        .Importance = olImportanceLow
        .Sensitivity = olNormal
      
        .Display
      
        emailBody = "<br>Happy " & Format(Now(), "dddd") & " All,<br><br> Please see below weekly activity update for the week starting " & Format(dtWeekStart + 1, "dddd mmmm dd, yyyy") & ":<br><br>"
        emailBody = emailBody & GetTaskReport_EmailBody(dtWeekStart)
        emailBody = emailBody & "Thanks!"
        'note: Displaying the email body (above) before setting the body value (below)
        '      should result in the user's default new-email signature being added
        .HTMLBody = emailBody & objMsg.HTMLBody
    End With
    
    Set objMsg = Nothing
End Sub

'Get list of active tasks
Private Function GetTaskReport_EmailBody(dtWeekStart As Date) As String
    On Error GoTo errOut
    
    Dim session As Outlook.NameSpace
    Dim taskFolder As Outlook.Folder
    Dim currentItem As Object
    Dim t As TaskItem
    Dim rs:  Set rs = CreateObject("ADODB.Recordset")
    'Dim rs As ADODB.Recordset
    Dim i As Integer
    Dim emailBody As String
    Dim lastCompany As String
    Dim lastRole As String
    Dim itemPrefix As String
    Dim qt As String: qt = "" 'placeholder, in case you want to add quotes (or other surrounding strings)
    Dim preHTML As String, postHTML As String
    
    
    'Build a recordset object, for ease of sorting and
    Set session = Application.session
    Set taskFolder = session.GetDefaultFolder(olFolderTasks)
    
    Set rs = New Recordset
    rs.Fields.Append "Ordinal", adBigInt
    rs.Fields.Append "Company", adVarChar, 100
    rs.Fields.Append "Role", adVarChar, 500
    rs.Fields.Append "Subject", adVarChar, 1000
    rs.Fields.Append "Status", adVarChar, 100
    rs.Fields.Append "IsRecurring", adSmallInt
    rs.Fields.Append "Priority", adVarChar, 50
    rs.Fields.Append "StartDate", adDate
    rs.Fields.Append "DueDate", adDate
    rs.Fields.Append "CompleteDate", adDate
    rs.Fields.Append "CreateTS", adDBTimeStamp
    rs.Fields.Append "OrderByValue", adVarChar, 500
    rs.CursorLocation = adUseClient
    rs.Open
     
    '''''''''''''''''
    'STEP 1: Loop thru all Task objects in Outlook, qualify the rows we want and add to our recordset object
    '''''''''''''''''
    For Each currentItem In taskFolder.Items
        If currentItem.Class = olTask Then
            Set t = currentItem
            
            If t.DateCompleted > dtWeekStart And t.Companies <> "Personal" Then
                
                rs.AddNew 'add TASK information to the recordset
                rs.Fields("Ordinal") = -t.Ordinal
                rs.Fields("Company") = t.Companies
                rs.Fields("Role") = t.Role
                rs.Fields("Subject") = t.Subject
                    If t.status = olTaskComplete Then rs.Fields("Status") = "Complete"
                    If t.status = olTaskDeferred Then rs.Fields("Status") = "Suspended"
                    If t.status = olTaskInProgress Then rs.Fields("Status") = "In Progress"
                    If t.status = olTaskNotStarted Then rs.Fields("Status") = "Not Started"
                    If t.status = olTaskWaiting Then rs.Fields("Status") = "Waiting"
                rs.Fields("IsRecurring") = t.IsRecurring
                    If t.Importance = olImportanceHigh Then rs.Fields("Priority") = "<font color=red>"
                    If t.Importance = olImportanceNormal Then rs.Fields("Priority") = ""
                    If t.Importance = olImportanceLow Then rs.Fields("Priority") = "<font color=gray>"
                rs.Fields("StartDate") = t.StartDate
                rs.Fields("DueDate") = t.DueDate
                rs.Fields("CompleteDate") = t.DateCompleted
                rs.Fields("CreateTS") = t.CreationTime
                
                If t.Role = "NOTE:" Then 'always sort NOTE at bottom of list
                    rs.Fields("OrderByValue") = "zzzzzzz"
                Else
                    rs.Fields("OrderByValue") = t.Role
                End If
            End If
        End If
    Next currentItem
    
    
    '''''''''''''''''
    'STEP 2: Iterate thru our qualified recordset object and build our HTML email body
    '''''''''''''''''
    rs.Sort = "Company asc, OrderByValue asc, IsRecurring, Status, CompleteDate asc"
    rs.MoveFirst
    
    lastCompany = ""
    lastRole = ""
    emailBody = "<ul>" & vbNewLine
        
    'loop thru recordset (containing only qualified rows)
    For i = 1 To rs.RecordCount
        
        If UCase(lastCompany) <> UCase(rs!Company) Then
            emailBody = emailBody & " <li><b><h3>" & qt & rs!Company & qt & "</h3></b></li>" & vbNewLine
            emailBody = emailBody & "   <ul> " & vbNewLine
            lastRole = "" 'always trigger a new role (aka effort) when changing companies
        End If

        'add Role Bullet Header, if Role has changed
        If UCase(lastRole) <> UCase(rs!Role) Then
            emailBody = emailBody & "     <li>" & qt & rs!Role & qt & "</li>" & vbNewLine
            emailBody = emailBody & "       <ul> " & vbNewLine
        End If
        
        
        'Detailed Task list (main bullet points):
        If rs!status = "Complete" Then 'Simplify "Complete" to "Done"
            itemPrefix = "Done: "
            preHTML = "<strike>"
            postHTML = "</strike>"
        Else
            itemPrefix = "ToDo: " 'Simplify everything that is NOT "Complete" as "ToDo"
            preHTML = ""
            postHTML = ""
        End If
        
        If rs!IsRecurring <> 0 Then itemPrefix = "Ongoing: " 'if Task is reoccurring, mark as "Ongoing"
        If Mid(rs!Subject, 1, 5) = "Note:" Then
            itemPrefix = "" 'something starting with "Note:" is not prefixed
            preHTML = "<i>"
            postHTML = "</i>"
        End If
        
        'if task is (a) not "Complete" and (b) not "Normal" priority, apply new font color
        If rs.Fields("Priority") <> "" And itemPrefix = "ToDo: " Then
            preHTML = preHTML & rs!Priority
            postHTML = "</font>" & postHTML
        End If
        
        emailBody = emailBody & "         <li>" & preHTML & qt & itemPrefix & rs!Subject & qt & postHTML & "</li>" & vbNewLine

        
        '''''increment''''''
        lastCompany = rs!Company
        lastRole = rs!Role
        rs.MoveNext
        If rs.EOF Then GoTo ExitLoop
        ''''''''''''''''''''
        
        
        'add Role Bullet Footer, if Role has changed
        If UCase(lastRole) <> UCase(rs!Role) Then
            emailBody = emailBody & "       </ul> " & vbNewLine
        End If
        
        'add Role Bullet Footer, if Role has changed
        If UCase(lastCompany) <> UCase(rs!Company) Then
            emailBody = emailBody & "   </ul> " & vbNewLine
        End If
        

    Next i
    
ExitLoop:
    emailBody = emailBody & "       </ul> " & vbNewLine
    emailBody = emailBody & "   </ul> " & vbNewLine
    emailBody = emailBody & "</ul> " & vbNewLine
        
    'return string to function
    GetTaskReport_EmailBody = emailBody

Exit Function
errOut:
    MsgBox "Holy Crap, there was an error!!!" & vbNewLine & "(" & Err.Number & ") " & Err.Description & vbNewLine & vbNewLine & "(um... good luck with that.)", vbCritical, "Error"
        
End Function
