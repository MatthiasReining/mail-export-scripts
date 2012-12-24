VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Outlook-Mail-Export"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
     
Public Sub startExport()
    Dim exportFolderRoot As String
    exportFolderRoot = "F:\My Mail\text-archive"
    
    For Each Folder In Application.GetNamespace("MAPI").Folders
        'MsgBox Folder
        Debug.Print "execute root folder: " + Folder
        'If (Folder = "Archivordner") Then
            processFolder Folder, exportFolderRoot
        'End If
    Next
        
End Sub


Private Sub processFolder(ByVal oParent As Outlook.MAPIFolder, parentFolder As String)

        Dim oFolder As Outlook.MAPIFolder
        'Dim oMail As Outlook.MailItem
        Dim oMail As Variant
        
        
        Dim folderName As String
        
        If (InStr(oParent.Name, "Gelöschte") > 0) Then
            Exit Sub
        End If
        
        
        
        folderName = parentFolder & "\" & StripIllegalChar(oParent.Name)
              
        Debug.Print "-> " & folderName
        
        If Dir(folderName, vbDirectory) = "" Then
            MkDir folderName
        End If
      
        Debug.Print "   " & oParent.Items.Count & "mails"
        
        For Each oMail In oParent.Items
            Dim mi As MailItem
            
            If oMail.Class = olMail Then
                Set mi = oMail
                'If InStr(mi.subject, "Ausbilderschulung CBU") > 0 Then
                    exportMail mi, folderName
                'End If
                
                
            End If

        Next

        Dim i As Integer
        i = oParent.Folders.Count
        
        If (i > 0) Then
        
            For Each oFolder In oParent.Folders
                processFolder oFolder, folderName
            Next
        End If
        
        Exit Sub

End Sub

Private Sub exportMail(objItem As MailItem, folderName As String)
    On Error Resume Next
    Dim fullfilename As String
    Dim filename As String
    Dim time As String
    Dim sender As String
    
    
    time = Format(objItem.ReceivedTime, "YYYY-MM-DD HHMMSS")
    sender = objItem.SenderName
    If (sender = "Matthias Reining") Then
        sender = objItem.To
    End If
    sender = StripIllegalChar(sender)
        
    filename = time & " " & StripIllegalChar(objItem.subject) & " (" & sender & ")"
    filename = Left(filename, 100) 'cut off, to long throws exception
        
    fullfilename = folderName & "\" & filename
    
    'If PathExists(fullfilename & ".txt") Then
    '    Exit Sub
    'End If
    
    Dim content As String
    content = ""
    content = content + "Von:        " & objItem.SenderName & "  (" & objItem.SenderEmailAddress & ")<br/>"
    content = content + "Gesendet:   " & Format(objItem.ReceivedTime, "Long Date") & " " & Format(objItem.ReceivedTime, "Short Time") & "<br/>"
    content = content + "An:         " & objItem.To & "<br/>"
    content = content + "Betreff:    " & objItem.subject & "<br/>"
    content = content + "<br/><br/>"
    content = content + objItem.HTMLBody
    
    
    Call txt_WriteAll(fullfilename & ".html", content)
    'objItem.SaveAs fullfilename & ".txt", olSaveAsText
    'objItem.SaveAs fullfilename & ".html", OlSaveAsType.olHTML

    Dim Atmt As Outlook.Attachment

    For Each Atmt In objItem.Attachments
        attFilename = fullfilename & "-" & Atmt.filename
        Atmt.SaveAsFile (attFilename)
    Next Atmt
    
End Sub

' Beliebigen Text in eine Textdatei speichern, wobei
' der bisherige Inhalt der Textdatei vollständig
' überschrieben wird
'
' sFilename:  vollständiger Dateiname
' sLines   :  Inhalt, der gespeichert werden soll
' ===============================================
Public Sub txt_WriteAll(ByVal sFilename As String, _
  ByVal sLines As String)
 
  Dim F As Integer
 
  ' Datei zum Schreiben öffnen
  ' Achtung: bisheriger Inhalt wird gelöscht!
  F = FreeFile
  Open sFilename For Output As #F
  Print #F, sLines
  Close #F
End Sub



Function StripIllegalChar(StrInput)
     
    Dim RegX            As Object
     
    Set RegX = CreateObject("vbscript.regexp")
     
    RegX.Pattern = "[\" & Chr(34) & "\!\#\$\%\^\&\*\(\)\=\+\|\[\]\{\}\`\'\;\:\<\>\?\/\,]"
    RegX.IgnoreCase = True
    RegX.Global = True
     
    StripIllegalChar = RegX.Replace(StrInput, "")
     
ExitFunction:
     
    Set RegX = Nothing
     
End Function

Private Function PathExists(ByVal strPath As String) As Boolean
  On Error Resume Next
  GetAttr strPath
  PathExists = (Err = 0)
End Function


