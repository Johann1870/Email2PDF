Attribute VB_Name = "MsgAttchmentHandler"
'Attribute VB_Name = "MessageAttachmentHandler"
'I believe this first came from BlueDevilFan http://www.experts-exchange.com/members/BlueDevilFan.html
'I've made my own modifications over the years.
'The macro takes a maximum of fourteen parameters.
'In the code below these are represented as P1 through P14.
'The parameters are positional (i.e. they must appear in the sequence given).

'P1.  Print the message.
    'Tells the macro to print the email.
    'Valid values are True or False.

'P2.  Save the message.
    'Tells the macro to save the email to the file system.
    'Valid values are True or False.

'P3.  Print the attachments.
    'Tells the macro to print the attachments.
    'Valid values are True or False.

'P4.  Save the attachments.
    'Tells the macro to save the attachments.
    'Valid values are True or False.

'P5.  Remove attachments.
    'Tells the macro to remove the attachments and insert hyperlinks to  them at the bottom of the message.
    'Valid values are True or False.

'P6. Attachment types.
    'Tells the macro what attachment types to save/print.
    'The macro will only process attachments that match the file types.
    'This parameter is a comma separated list of file extensions.
    'For example, to only process Word documents (both 2007 and earlier) you’d set this parameter to "doc,docx".

'P7. Target file system folder.
    'This tells the macro which file system folder to save the message and/or attachments to.
    'Valid values are any existing file system folder, including network shares and UNC paths.
    'If you’ve told the macro to save the message and/or attachments and you fail to specify this parameter,
    'then the macro will save the items to your My Documents folder.

'P8. Message save format.
    'Tells the macro what format to save the message in (assuming that you are saving the message).
    'You will be prompted with a list of valid values when you enter this parameter.

'P9. Printer name.
    'Tells the macro what printer to print the message and/or attachments to.
    'This allows you to print to any printer, not just the default printer.
    'Valid values include the name of any printer that appears in your list of printers.
    'If you’ve told the macro to print the message and/or attachments and you don’t specify a printer,
    'then the macro will print them to your default printer.
    
'P10. File ID
    'This is the name of the file
    'When P13 is turned on, the name appears after the date
    'Ex. 02 October 2012 "Production Report".pdf
    'If both P10 and P13 are left blank, the file will be saved as the name of the original attachment.
    
'P11. File Extension
    'This is the file extension the file will be saved under.
    'It is important to include the period
    'Ex. ".pdf"
    
'P12. Date Folder Hierarchy
    'Tells the macro to append a date hierarchy to the folder path
    ' i.e., Path...\2012\October\
    'Valid values are True or False
    
'P13. Date in File Name
    'Tells the macro to add the date to the beginning of the file name
    ' i.e., 02 October 2012 Production Report.pdf
    'Works with P10
    'If both P10 and P13 are left blank, the file will be saved as the name of the original attachment.
    'Valid values are True or False
    
'P14. Create Copy or Overwrite
    'This determines what happens in the event of a file already existing with the same name in the folder path specified.
    'If set to true, the file will create a copy: Copy (1) of FileName.pdf
    'If set to false, the file will overwrite the existing file.
    

'All the parameters are optional so you can omit those that you don’t need.


Sub RecipeSuggestion(Item As Outlook.MailItem)
        'Recipe Suggestions
        MessageAndAttachmentProcessor Item, False, True, False, False, False, , "J:\rec\suggestions", olDoc, , "Recipe Suggestion ", ".doc", False, False, False
        
        Dim objShell As Object
        Set objShell = CreateObject("WScript.Shell")
        objShell.Run ("cscript J:\sc\vbcallcmd.vbs")
        Set objShell = Nothing
End Sub
