'Attribute VB_Name = "MessageAttachmentProcessor"
'I believe this first came from BlueDevilFan http://www.experts-exchange.com/members/BlueDevilFan.html
'I've made my own modifications over the years.

Public Declare Function GetProfileString Lib "kernel32" Alias "GetProfileStringA" _
        (ByVal lpAppName As String, ByVal lpKeyName As String, _
        ByVal lpDefault As String, ByVal lpReturnedString As String, _
        ByVal nSize As Long) As Long

Private Declare Function ShellExecute Lib "shell32.dll" _
  Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, _
  ByVal lpFile As String, ByVal lpParameters As String, _
  ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Sub TestMacro()
    MsgBox "Macros are enabled."
End Sub


Sub MessageAndAttachmentProcessor(Item As Outlook.MailItem, _
    Optional bolPrintMsg As Boolean, _
    Optional bolSaveMsg As Boolean, _
    Optional bolPrintAtt As Boolean, _
    Optional bolSaveAtt As Boolean, _
    Optional bolInsertLink As Boolean, _
    Optional strAttFileTypes As String, _
    Optional strFolderPath As String, _
    Optional varMsgFormat As OlSaveAsType, _
    Optional strPrinter As String, _
    Optional strFileID As String, _
    Optional strSaveAsFileExt As String, _
    Optional bolDtFolderHeir As Boolean, _
    Optional bolDtFileName As Boolean, _
    Optional bolDuporOverwrite As Boolean)
   
    Dim olkAttachment As Outlook.Attachment, _
        objFSO As Object, _
        fso, _
        strMyPath As String, _
        strExtension As String, _
        strFileName As String, _
        strOriginalPrinter As String, _
        strLinkText As String, _
        strRootFolder As String, _
        strTempFolder As String, _
        varFileType As Variant, _
        intCount As Integer, _
        intIndex As Integer, _
        arrFileTypes As Variant

    Set objFSO = CreateObject("Scripting.FileSystemObject")
    strTempFolder = Environ("TEMP") & "\"
    
    
    
    
    If strAttFileTypes = "" Then
        arrFileTypes = Array("*")
    Else
        arrFileTypes = Split(strAttFileTypes, ",")
    End If
    
    If bolPrintMsg Or bolPrintAtt Then
        If strPrinter <> "" Then
            strOriginalPrinter = GetDefaultPrinter()
            SetDefaultPrinter strPrinter
        End If
    End If
    
    
    
'Saving Messages or Attachments
    If bolSaveMsg Or bolSaveAtt Then
        
        'If Parameter is empty set as default to "My Documents"
        If strFolderPath = "" Then
            strRootFolder = Environ("USERPROFILE") & "\My Documents\"
        Else
            'Makes sure path has backslash at the end
            strRootFolder = strFolderPath & IIf(Right(strFolderPath, 1) = "\", "", "\")
        End If
        
        If bolDtFolderHeir Then
            'Adds year folder to path
                Set fso = CreateObject("Scripting.FileSystemObject")
                strRootFolder = strRootFolder & Year(Now)
                
                If Not fso.FolderExists(strRootFolder) Then
                    fso.CreateFolder (strRootFolder)
                End If
            
            'Adds month folder to path
                strRootFolder = strRootFolder & "\" & Format(Now, "mmmm") & "\"
                
                If Not fso.FolderExists(strRootFolder) Then
                    fso.CreateFolder (strRootFolder)
                End If
        End If
    End If
    
    
    
    
    
    
    
    If bolSaveMsg Then
        Select Case varMsgFormat
            Case olHTML
                strExtension = ".html"
            Case olMSG
                strExtension = ".msg"
            Case olRTF
                strExtension = ".rtf"
            Case olDoc
                strExtension = ".doc"
            Case olTXT
                strExtension = ".txt"
            Case Else
                strExtension = ".msg"
        End Select
        Item.SaveAs strRootFolder & RemoveIllegalCharacters(Item.Subject) & strExtension, IIf(varMsgFormat <> 0, varMsgFormat, olMSG)
    End If
        
    For intIndex = Item.Attachments.Count To 1 Step -1
        Set olkAttachment = Item.Attachments.Item(intIndex)
        
'Print the attachments if requested'
        If bolPrintAtt Then
            If olkAttachment.Type <> olEmbeddeditem Then
                For Each strFileType In arrFileTypes
                    If (strFileType = "*") Or (LCase(objFSO.GetExtensionName(olkAttachment.FileName)) = LCase(strFileType)) Then
                        olkAttachment.SaveAsFile strTempFolder & olkAttachment.FileName
                        ShellExecute 0&, "print", strTempFolder & olkAttachment.FileName, 0&, 0&, 0&
                    End If
                Next
            End If
        End If
        
        
 
        
'Save the attachments if requested'
        If bolSaveAtt Then
            
            If strFileID = "" Then
                If bolDtFileName Then
                    strFileName = Format(Now, "dd mmmm yyyy") & strSaveAsFileExt
                Else
                    strFileName = olkAttachment.FileName
                End If
            Else
                
                If bolDtFileName Then
                    strFileName = Format(Now, "dd mmmm yyyy") & " " & strFileID & strSaveAsFileExt
                Else
                    strFileName = strFileID & strSaveAsFileExt
                End If
                
            End If
            
            
            
            intCount = 0
            Do While True
                strMyPath = strRootFolder & strFileName
                
    'Duplicate or Overwrite
                If objFSO.FileExists(strMyPath) Then
                    
                    If bolDuporOverwrite Then
                        intCount = intCount + 1
                        strFileName = "Copy (" & intCount & ") of " & olkAttachment.FileName
                    Else
                        SetAttr strMyPath, vbNormal
                        Kill strMyPath
                    End If
                
                Else
                    Exit Do
                End If
            Loop
            olkAttachment.SaveAsFile strMyPath
            If bolInsertLink Then
                If Item.BodyFormat = olFormatHTML Then
                    strLinkText = strLinkText & "<a href=""file://" & strMyPath & """>" & olkAttachment.FileName & "</a><br>"
                Else
                    strLinkText = strLinkText & strMyPath & vbCrLf
                End If
                olkAttachment.Delete
            End If
        End If
    Next
    
    

    
    
    If bolPrintMsg Then
        Item.PrintOut
    End If
    
    If bolPrintMsg Or bolPrintAtt Then
        If strOriginalPrinter <> "" Then
            SetDefaultPrinter strOriginalPrinter
        End If
    End If
    
    If bolInsertLink Then
        If Item.BodyFormat = olFormatHTML Then
            Item.HTMLBody = Item.HTMLBody & "<br><br>Removed Attachments<br><br>" & strLinkText
        Else
            Item.Body = Item.Body & vbCrLf & vbCrLf & "Removed Attachments" & vbCrLf & vbCrLf & strLinkText
        End If
        Item.Save
    End If

    Set objFSO = Nothing
    Set olkAttachment = Nothing
End Sub

Function GetDefaultPrinter() As String
    Dim strPrinter As String, _
        intReturn As Integer
    strPrinter = Space(255)
    intReturn = GetProfileString("Windows", ByVal "device", "", strPrinter, Len(strPrinter))
    If intReturn Then
        strPrinter = UCase(Left(strPrinter, InStr(strPrinter, ",") - 1))
    End If
    GetDefaultPrinter = strPrinter
End Function

Function RemoveIllegalCharacters(strValue As String) As String
    ' Purpose: Remove characters that cannot be in a filename from a string.'
    ' Written: 4/24/2009'
    ' Author:  BlueDevilFan'
    ' Outlook: All versions'
    RemoveIllegalCharacters = strValue
    RemoveIllegalCharacters = Replace(RemoveIllegalCharacters, "<", "")
    RemoveIllegalCharacters = Replace(RemoveIllegalCharacters, ">", "")
    RemoveIllegalCharacters = Replace(RemoveIllegalCharacters, ":", "")
    RemoveIllegalCharacters = Replace(RemoveIllegalCharacters, Chr(34), "'")
    RemoveIllegalCharacters = Replace(RemoveIllegalCharacters, "/", "")
    RemoveIllegalCharacters = Replace(RemoveIllegalCharacters, "\", "")
    RemoveIllegalCharacters = Replace(RemoveIllegalCharacters, "|", "")
    RemoveIllegalCharacters = Replace(RemoveIllegalCharacters, "?", "")
    RemoveIllegalCharacters = Replace(RemoveIllegalCharacters, "*", "")
End Function

Sub SetDefaultPrinter(strPrinterName As String)
    Dim objNet As Object
    Set objNet = CreateObject("Wscript.Network")
    objNet.SetDefaultPrinter strPrinterName
    Set objNet = Nothing
End Sub
