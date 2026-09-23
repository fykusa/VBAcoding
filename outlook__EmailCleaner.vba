' ============================================================
' Outlook VBA - Mazani duplicitnich mailu v Inboxu
' Klic duplicity: SenderEmailAddress + Subject, ReceivedTime se porovnava
' v casovem okne (viz DUPLICATE_WINDOW_MINUTES), ne na presnou minutu
' Nejprve zobrazi preview, pak ceka na potvrzeni.
' ============================================================
Option Explicit

' Casove okno (v minutach) pro povazovani mailu se stejnym
' odesilatelem+predmetem za duplicitu. Reseni pripadu, kdy dva
' stejne maily prijdou tesne kolem hranice minuty (napr. 13:00:59
' a 13:01:01) - presna shoda na minutu by je nespojila.
Public Const DUPLICATE_WINDOW_MINUTES As Long = 5

Public Sub RemoveInboxDuplicates()
    Dim inbox As Outlook.MAPIFolder
    Set inbox = Nothing

    Dim store As Outlook.store
    For Each store In Application.Session.Stores
        If store.ExchangeStoreType = olPrimaryExchangeMailbox Then
            Set inbox = store.GetDefaultFolder(olFolderInbox)
            Exit For
        End If
    Next store

    If inbox Is Nothing Then
        MsgBox "Inbox nenalezen.", vbExclamation
        Exit Sub
    End If

    Dim total As Long
    total = inbox.Items.count

    frmProgress.Caption = "Hledam duplicity v Inboxu"
    frmProgress.lblStatus.Caption = "Nacitam maily... (celkem " & total & ")"
    frmProgress.lblCounter.Caption = ""
    frmProgress.txtResults.Visible = False
    frmProgress.btnClose.Visible = False
    frmProgress.Show vbModeless
    DoEvents

    ' Scripting.Dictionary pro O(1) lookup
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")

    Dim toDeleteId()    As String
    Dim toDeleteDesc()  As String   ' popis pro preview
    Dim delCount        As Long
    delCount = 0
    ReDim toDeleteId(total)
    ReDim toDeleteDesc(total)

    Dim itm  As Object
    Dim i    As Long
    i = 0

    For Each itm In inbox.Items
        i = i + 1
        If i Mod 20 = 0 Then
            frmProgress.lblStatus.Caption = "Prochazim " & i & " / " & total
            frmProgress.lblCounter.Caption = "Duplicity: " & delCount
            DoEvents
        End If

        Dim s As String: s = ""
        Dim u As String: u = ""
        Dim t As Date: t = 0
        Dim key As String: key = ""

        On Error Resume Next
        s = itm.SenderEmailAddress
        u = itm.subject
        t = itm.ReceivedTime
        On Error GoTo 0

        ' Klic BEZ casu - cas se resi zvlast pres casove okno,
        ' aby se nerozbily duplicity tesne kolem hranice minuty.
        key = LCase(Trim(s)) & "|" & LCase(Trim(u))

        If Len(key) < 3 Then GoTo NextItem

        Dim isDuplicate As Boolean
        isDuplicate = False

        If dict.Exists(key) Then
            Dim seenTimes As Collection
            Set seenTimes = dict(key)

            Dim k As Long
            For k = 1 To seenTimes.count
                If Abs(DateDiff("n", seenTimes(k), t)) <= DUPLICATE_WINDOW_MINUTES Then
                    isDuplicate = True
                    Exit For
                End If
            Next k

            If Not isDuplicate Then seenTimes.Add t
        Else
            Set seenTimes = New Collection
            seenTimes.Add t
            dict.Add key, seenTimes
        End If

        If isDuplicate Then
            On Error Resume Next
            toDeleteId(delCount) = itm.EntryID
            toDeleteDesc(delCount) = Format(t, "DD.MM HH:MM") & "  " & _
                                     Left(Trim(s), 30) & "  |  " & Left(Trim(u), 50)
            On Error GoTo 0
            delCount = delCount + 1
        End If

NextItem:
        Set itm = Nothing
    Next itm

    Set dict = Nothing

    ' --- Preview ---
    Dim preview As String
    preview = "NALEZENE DUPLICITY (" & delCount & ")" & vbCrLf
    preview = preview & String(70, "-") & vbCrLf
    preview = preview & PadR("Datum", 12) & "  " & PadR("Odesilatel", 30) & "  Predmet" & vbCrLf
    preview = preview & String(70, "-") & vbCrLf

    Dim d As Long
    For d = 0 To delCount - 1
        preview = preview & toDeleteDesc(d) & vbCrLf
    Next d

    If delCount = 0 Then
        preview = preview & "(zadne duplicity nenalezeny)" & vbCrLf
    End If

    preview = preview & vbCrLf & "Klikni OK v nasledujicim dialogu pro presun do Odstranenych," & vbCrLf
    preview = preview & "nebo Zrusit pro ukonceni bez smazani."

    With frmProgress
        .Caption = "Duplicity – preview"
        .lblStatus.Caption = "Nalezeno duplicit: " & delCount & " (z " & total & " mailu)"
        .lblCounter.Caption = "Zkontroluj seznam nize a potvrdte mazani."
        .txtResults.Text = preview
        .txtResults.Visible = True
        .btnClose.Visible = True
    End With

    If delCount = 0 Then Exit Sub

    ' Cekej na potvrzeni
    Dim ans As Integer
    ans = MsgBox("Presunout " & delCount & " duplicit do Odstranenych polozek?", _
                 vbOKCancel + vbQuestion, "Potvrdit mazani")
    If ans <> vbOK Then
        frmProgress.lblStatus.Caption = "Zruseno, nic nebylo smazano."
        Exit Sub
    End If

    ' --- Mazani ---
    frmProgress.lblStatus.Caption = "Presuvam duplicity..."
    frmProgress.lblCounter.Caption = ""
    frmProgress.txtResults.Visible = False
    DoEvents

    Dim deleted As Long
    deleted = 0
    For d = 0 To delCount - 1
        On Error Resume Next
        Dim delItm As Object
        Set delItm = Application.Session.GetItemFromID(toDeleteId(d))
        If Not delItm Is Nothing Then
            delItm.Delete
            deleted = deleted + 1
        End If
        Set delItm = Nothing
        Err.Clear
        On Error GoTo 0
        If d Mod 10 = 0 Then DoEvents
    Next d

    ' --- Vysledek ---
    Dim msg As String
    msg = "VYSLEDEK CISTENI INBOXU" & vbCrLf
    msg = msg & String(40, "-") & vbCrLf
    msg = msg & "Celkem mailu v Inboxu:     " & total & vbCrLf
    msg = msg & "Duplicit nalezeno:         " & delCount & vbCrLf
    msg = msg & "Presunuto do Odstranenych: " & deleted & vbCrLf
    msg = msg & String(40, "-") & vbCrLf & vbCrLf
    msg = msg & "Zkontroluj Odstranene polozky a pak vyprazdni kos."

    With frmProgress
        .Caption = "Cisteni Inboxu – hotovo"
        .lblStatus.Caption = "Hotovo. Presunuto: " & deleted & " duplicit."
        .lblCounter.Caption = ""
        .txtResults.Text = msg
        .txtResults.Visible = True
    End With
End Sub

Private Function PadR(s As String, width As Integer) As String
    PadR = Left(s & Space(width), width)
End Function


