' ============================================================
' Outlook VBA - some tools
'
' ============================================================


Sub CopySelectedMailToChosenFolder()
    Dim objItem As Object
    Dim objMail As MailItem
    Dim objTargetFolder As Outlook.Folder
    
    ' Otevře dialog pro výběr složky
    Set objTargetFolder = Application.Session.PickFolder
    
    ' Pokud uživatel zruší výběr, ukončí makro
    If objTargetFolder Is Nothing Then
        MsgBox "Výběr složky byl zrušen.", vbExclamation
        Exit Sub
    End If
    
    ' Projde všechny označené položky
    For Each objItem In Application.ActiveExplorer.Selection
        If TypeName(objItem) = "MailItem" Then
            Set objMail = objItem
            objMail.Copy.Move objTargetFolder
        End If
    Next
    
    MsgBox "E-mail(y) byly zkopírovány do složky: " & objTargetFolder.name, vbInformation
End Sub



