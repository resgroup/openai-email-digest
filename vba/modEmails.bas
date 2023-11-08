Attribute VB_Name = "modEmails"
Option Explicit

Private Const M_STR_EMAIL_FOLDER As String = "C:\Emails\"


Public Sub RunSaveEmailsToJsonFiles()

    Static lastsaved As Date
    Dim lngSecondsSinceLastRun As Long
    
    If lastsaved > 0 Then
        lngSecondsSinceLastRun = DateDiff("s", lastsaved, Now)
    Else
        lngSecondsSinceLastRun = -1
    End If
    
    If lngSecondsSinceLastRun > 600 Or lngSecondsSinceLastRun < 0 Then
        SaveEmailsToJsonFiles
        lastsaved = Now
    End If

End Sub


Function FormatDate(intpuDate As Date) As String
    FormatDate = Format(intpuDate, "yyyy-mm-dd")
End Function

Function GetFolderPath(strdate As String) As String
    GetFolderPath = M_STR_EMAIL_FOLDER & strdate & "\"
End Function

Function GetEmailsForDate(myInbox As Outlook.Folder, strdate As String) As Outlook.Items
    
    Dim filter As String
    ' Define a filter to get only today's emails
    filter = "[ReceivedTime] >= '" & strdate & " 00:00 AM' AND [ReceivedTime] <= '" & strdate & " 11:59 PM'"
    
    ' Apply the filter to the Inbox Items
    Set GetEmailsForDate = myInbox.Items.Restrict(filter)
    
End Function

Sub SaveEmailsToJsonFiles()
    Dim myNamespace As Outlook.NameSpace
    Dim myInbox As Outlook.Folder
    Dim myItems As Outlook.Items
    Dim myItem As Object
    Dim fso As Object
    Dim jsonFile As Object
    Dim folderPath As String
    Dim fileName As String
    Dim currentDate As String
    Dim lngCount As Long
    Dim filePath As String
    Dim emailSubject As String
    Dim emailBody As String
    Dim jsonContent As String
    Dim i As Integer
    
    ' Create a folder path with the current date
    currentDate = FormatDate(Now)
    folderPath = GetFolderPath(currentDate)
    
    ' Set Namespace and Folder objects
    Set myNamespace = Application.GetNamespace("MAPI")
    Set myInbox = myNamespace.GetDefaultFolder(olFolderInbox) ' Refer to the Inbox
    
    ' Clean Up Emails over last 3 days
    For i = 0 To 3
        CleanupEmailJsonFiles myInbox, FormatDate(DateAdd("d", -i, Now))
    Next

    Set myItems = GetEmailsForDate(myInbox, currentDate)
    
    ' Create FileSystemObject
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Create a folder for today's date if it doesn't exist
    If Not fso.FolderExists(folderPath) Then
        fso.CreateFolder folderPath
    End If
    
    For Each myItem In myItems
        
        If TypeName(myItem) = "MailItem" Then
        
            ' Use EntryID as the filename
            fileName = myItem.EntryID
            filePath = folderPath & fileName & ".json"
            
            ' Write email to file if it doesnt already exist
            If Dir(filePath) = "" Then
                           
                On Error Resume Next
                emailSubject = myItem.Subject
                If Err.Number <> 0 Then
                    MsgBox (Len(myItem.Subject))
                    emailSubject = "Subject could not be retrieved."
                    Err.Clear
                End If
                On Error GoTo 0
                
                
                emailBody = myItem.Body
                If Err.Number <> 0 Then
                    MsgBox (Len(myItem.Body))
                    emailBody = "Body could not be retrieved."
                    Err.Clear
                End If
                On Error GoTo 0
                
                
                jsonContent = "{"
                jsonContent = jsonContent & """Subject"": """ & JsonStringEscape(emailSubject) & ""","
                jsonContent = jsonContent & """SenderName"": """ & JsonStringEscape(myItem.SenderName) & ""","
                jsonContent = jsonContent & """SenderEmailAddress"": """ & JsonStringEscape(myItem.SenderEmailAddress) & ""","
                jsonContent = jsonContent & """ReceivedTime"": """ & myItem.receivedTime & ""","
                jsonContent = jsonContent & """Body"": """ & JsonStringEscape(emailBody) & """"
                jsonContent = jsonContent & "}"
                
                ' Write the email to a JSON file
                
                'SetClipboard jsonContent
                WriteTextFileUTF8 folderPath & fileName & ".json", jsonContent
                
            End If
        End If
        
    Next myItem
    
    SaveLastRunDateToJsonFile
    
    ' Clear memory
    Set fso = Nothing
    Set myItem = Nothing
    Set myItems = Nothing
    Set myInbox = Nothing
    Set myNamespace = Nothing
    
    
End Sub


Function JsonStringEscape(str As String) As String

    Dim result As String
    Dim i As Long
    Dim ch As String
    
    ' Iterate through each character in the string
    For i = 1 To Len(str)
        ch = Mid$(str, i, 1)
        
        ' Replace special JSON characters with escaped versions
        Select Case ch
            Case "\"""
                result = result & "\\"
            Case """"
                result = result & "\"""
            Case "/"
                result = result & "\/"
            Case vbBack
                result = result & "\b"
            Case vbFormFeed
                result = result & "\f"
            Case vbLf
                result = result & "\n"
            Case vbCr
                result = result & "\r"
            Case vbTab
                result = result & "\t"
            Case """"
                result = result & "\"""
            Case Else
                ' Append other characters directly
                result = result & ch
        End Select
    Next i
    
    ' Assign escaped string to the function return
    JsonStringEscape = result
    
End Function

Sub WriteTextFileUTF8(filePath As String, text As String)
    Dim stream As Object
    
    If Dir(filePath) <> "" Then
        ' File has already been written
        Exit Sub
    End If
    
    ' Create a new ADODB stream
    Set stream = CreateObject("ADODB.Stream")
    
    ' Define stream type - we want to save text/string data.
    stream.Type = 2 'Specify text data
    ' Specify charset for the source text (we want to save it as UTF-8)
    stream.Charset = "utf-8"
    ' Open the stream And write text data
    stream.Open
    stream.WriteText text
    ' Save the content of a Stream object to a file
    stream.SaveToFile filePath, 2 ' 2 for adSaveCreateOverWrite
    
    ' Close the stream
    stream.Close
    
    ' Release the stream object
    Set stream = Nothing
End Sub


Sub CleanupEmailJsonFiles(myInbox As Outlook.Folder, ByVal targetDate As String)
    Dim fso As Object
    Dim inboxFolder As Outlook.Folder
    Dim mailItem As Outlook.mailItem
    Dim jsonFolderPath As String
    Dim currentEntryID As String
    Dim file As Object
    Dim myItems As Outlook.Items
    
    ' Create a dictionary to store the current emails' EntryIDs
    Dim dictCurrentEmails As Object
    Set dictCurrentEmails = CreateObject("Scripting.Dictionary")

    ' Set the folder path based on the target date
    jsonFolderPath = M_STR_EMAIL_FOLDER & "\" & targetDate & "\"
    
    ' Create FileSystemObject
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Only proceed if the folder exists
    If Not fso.FolderExists(jsonFolderPath) Then Exit Sub
    
    ' Get the Inbox folder
    Set myItems = GetEmailsForDate(myInbox, targetDate)
    
    ' Add all current emails' EntryID to the dictionary
    For Each mailItem In myItems
        If TypeName(mailItem) = "MailItem" Then
            ' Add the EntryID to the dictionary
            dictCurrentEmails(mailItem.EntryID) = True
        End If
    Next mailItem
    
    ' Loop through all files in the folderPath and delete any that do not have a matching EntryID in the inbox
    For Each file In fso.GetFolder(jsonFolderPath).Files
        If fso.GetExtensionName(file.Name) = "json" Then
            currentEntryID = fso.GetBaseName(file.Name)
            
            ' Check if the EntryID of the file is not in the dictionary
            If Not dictCurrentEmails.Exists(currentEntryID) Then
                file.Delete ' If not found in the dictionary, delete the file
            End If
        End If
    Next file


    ' Cleanup
    Set myItems = Nothing
    Set fso = Nothing
    Set dictCurrentEmails = Nothing
    
End Sub

Sub SaveLastRunDateToJsonFile()
    Dim fso As Object
    Dim jsonFile As Object
    Dim lastRunDate As String
    Dim filePath As String
    Dim jsonText As String
    
    ' Create an instance of the FileSystemObject
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Define the file path
    filePath = M_STR_EMAIL_FOLDER & "last_run.json"
    
    ' Get the current date and time
    lastRunDate = Format(Now, "yyyy-mm-dd HH:MM:ss")
    
    ' Create the JSON string
    jsonText = "{""lastRunDate"": """ & lastRunDate & """}"
    
    ' Check if the folder exists, if not create it
    If Not fso.FolderExists(M_STR_EMAIL_FOLDER) Then
        fso.CreateFolder M_STR_EMAIL_FOLDER
    End If
    
    ' Create a new file, or open the existing one
    Set jsonFile = fso.CreateTextFile(filePath, True)
    
    ' Write the JSON string to the file
    jsonFile.WriteLine jsonText
    
    ' Close the file
    jsonFile.Close
    
    ' Clear the memory
    Set jsonFile = Nothing
    Set fso = Nothing
    

End Sub

