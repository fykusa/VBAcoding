' ============================================================
' TENTO KOD PATRI DO SPECIALNIHO MODULU "ThisOutlookSession"
' (v Project Exploreru ve VBA editoru), NE do noveho modulu.
' Zkopiruj ho tam rucne - ThisOutlookSession uz obvykle neco
' obsahuje, tohle jen pridej k existujicimu obsahu.
'
' Vyzaduje outlook__EmailCleanerAuto.vba (Sub CheckAndRemoveDuplicateSilent)
' nahrany jako bezny modul ve stejnem projektu.
' ============================================================
Option Explicit

Private WithEvents InboxItems As Outlook.Items

Private Sub Application_Startup()
    InitDuplicateWatcher
End Sub

' Verejna, aby slo watcher nahodit i rucne (Alt+F8) bez restartu Outlooku,
' napr. hned po vlozeni tohoto kodu.
Public Sub InitDuplicateWatcher()
    Dim inbox As Outlook.MAPIFolder
    Dim store As Outlook.store

    For Each store In Application.Session.Stores
        If store.ExchangeStoreType = olPrimaryExchangeMailbox Then
            Set inbox = store.GetDefaultFolder(olFolderInbox)
            Exit For
        End If
    Next store

    If Not inbox Is Nothing Then
        Set InboxItems = inbox.Items
        Debug.Print "DuplicateWatcher aktivni na: " & inbox.FolderPath
        LogWatcherStart inbox.FolderPath
    Else
        Debug.Print "DuplicateWatcher: Inbox nenalezen, watcher NENI aktivni."
        LogWatcherStart "CHYBA - Inbox nenalezen, watcher NENI aktivni"
    End If

    ' Uklid archivu starych .msg souboru (viz outlook__EmailCleanerAuto.vba)
    PurgeOldArchivedEmails
End Sub

Private Sub InboxItems_ItemAdd(ByVal Item As Object)
    ' Kazdy novy mail v Inboxu prochazi tudy.
    CheckAndRemoveDuplicateSilent Item
End Sub
