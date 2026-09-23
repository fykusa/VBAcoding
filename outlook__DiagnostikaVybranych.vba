' ============================================================
' Outlook VBA - Diagnostika oznacenych mailu
' Vypise SenderEmailAddress, SenderName, Subject a ReceivedTime
' (na vteriny) pro vsechny oznacene polozky a vysledek zkopiruje
' rovnou do schranky (Ctrl+V).
' ============================================================
Option Explicit

Public Sub DiagnostikaVybranych()
    Dim itm As Object
    Dim msg As String
    msg = ""

    If Application.ActiveExplorer.Selection.count = 0 Then
        MsgBox "Nejdriv oznac v Inboxu jeden nebo vice mailu.", vbExclamation
        Exit Sub
    End If

    For Each itm In Application.ActiveExplorer.Selection
        msg = msg & "EntryID: " & itm.EntryID & vbCrLf
        msg = msg & "SenderEmailAddress: [" & itm.SenderEmailAddress & "]" & vbCrLf
        msg = msg & "SenderName: [" & itm.SenderName & "]" & vbCrLf
        msg = msg & "Subject: [" & itm.subject & "]" & vbCrLf
        msg = msg & "ReceivedTime: " & Format(itm.ReceivedTime, "YYYY-MM-DD HH:NN:SS") & vbCrLf
        msg = msg & String(40, "-") & vbCrLf
    Next itm

    Debug.Print msg

    ' --- Kopie do schranky (late-bound, nevyzaduje pridani reference) ---
    Dim clipObj As Object
    On Error Resume Next
    Set clipObj = CreateObject("MSForms.DataObject")
    On Error GoTo 0

    If Not clipObj Is Nothing Then
        clipObj.SetText msg
        clipObj.PutInClipboard
        MsgBox "Zkopirovano do schranky (" & Application.ActiveExplorer.Selection.count & _
               " polozek). Vloz pomoci Ctrl+V.", vbInformation
    Else
        MsgBox "Nepodarilo se pristoupit ke schrance (MSForms.DataObject)." & vbCrLf & _
               "Vypis alespon v Immediate okne (Ctrl+G v editoru):" & vbCrLf & vbCrLf & msg, vbExclamation
    End If
End Sub
