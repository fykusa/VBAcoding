' ============================================================
' Outlook VBA - Ticha (autonomni) detekce duplicit pri prichodu mailu
'
' Pouziti:
'   1. Tento modul vloz do VBA projektu (staci jako novy standardni modul).
'   2. Do ThisOutlookSession vloz kod z
'      outlook__ThisOutlookSession_DuplicateWatcher.vba
'      (to je specialni modul, ktery nejde nahrat souborem - musi se
'      zkopirovat rucne do existujiciho ThisOutlookSession).
'   3. Restartuj Outlook (Application_Startup se spusti jen pri startu),
'      nebo rucne spust InitDuplicateWatcher z ThisOutlookSession.
'
' Chovani:
'   - Kdyz prijde novy mail do Inboxu, zkontroluje se, jestli uz v Inboxu
'     neexistuje jiny mail se stejnym odesilatelem + predmetem a casem
'     v okne DUPLICATE_WINDOW_MINUTES.
'   - Pokud ano, NOVY prichozi mail se nejdriv ulozi jako .msg do archivni
'     slozky (ARCHIVE_FOLDER_PATH) a pak se presune do Odstranenych
'     (puvodni/starsi kopie v Inboxu zustava), bez jakehokoliv dialogu.
'   - Kazda takova akce se zaloguje do textoveho souboru (LOG_FILE_PATH)
'     pro zpetnou kontrolu.
'   - Pri kazdem startu Outlooku (viz InitDuplicateWatcher v
'     ThisOutlookSession) se z archivni slozky smazou soubory starsi
'     nez ARCHIVE_RETENTION_DAYS dni.
' ============================================================
Option Explicit

Public Const AUTO_DUPLICATE_WINDOW_MINUTES As Long = 5
Public Const LOG_FILE_PATH As String = "C:\Users\FG408DJ\EmailCleanerAuto_log.txt"
Public Const ARCHIVE_FOLDER_PATH As String = "C:\Users\FG408DJ\OutlookDeletedArchive\"
Public Const ARCHIVE_RETENTION_DAYS As Long = 14

' Docasne zapnuto kvuli ladeni - loguje UPLNE KAZDE spusteni ItemAdd,
' i kdyz zadna duplicita nebyla nalezena. Az bude jasne, ze ItemAdd
' spolehlive stridli, vypni (False), at log neroste zbytecne rychle.
Public Const VERBOSE_LOGGING As Boolean = True

' Volano z ThisOutlookSession.InboxItems_ItemAdd
Public Sub CheckAndRemoveDuplicateSilent(ByVal newItem As Object)
    If TypeName(newItem) <> "MailItem" Then Exit Sub

    On Error GoTo ErrHandler

    Dim s As String, sn As String, u As String, t As Date
    s = ""
    sn = ""
    u = ""
    t = 0

    On Error Resume Next
    s = newItem.SenderEmailAddress
    sn = newItem.SenderName
    u = newItem.subject
    t = newItem.ReceivedTime
    On Error GoTo ErrHandler

    If VERBOSE_LOGGING Then
        LogLine "ADD | " & sn & " | " & u & " | " & Format(t, "YYYY-MM-DD HH:NN:SS")
    End If

    Dim key As String
    key = LCase(Trim(s)) & "|" & LCase(Trim(u))
    If Len(key) < 3 Then Exit Sub

    Dim newEntryId As String
    newEntryId = newItem.EntryID

    Dim inbox As Outlook.MAPIFolder
    Set inbox = GetPrimaryInbox()
    If inbox Is Nothing Then Exit Sub

    ' Prime prochazeni polozek serazenych podle casu prijeti (sestupne),
    ' NE Items.Restrict - ten jede pres server-side vyhledavaci index,
    ' ktery u prave prichozich polozek muze mit zpozdeni (zvlast kdyz
    ' dorazi vice mailu ve stejne synchronizacni davce) a novou
    ' duplicitu tak nemusi jeste "videt".
    Dim allItems As Outlook.Items
    Set allItems = inbox.Items
    On Error Resume Next
    allItems.Sort "[ReceivedTime]", True
    On Error GoTo ErrHandler

    Dim cand As Object
    Dim matchFound As Boolean
    matchFound = False

    Dim scanned As Long
    scanned = 0

    For Each cand In allItems
        scanned = scanned + 1
        If scanned > 300 Then Exit For   ' pojistka proti zbytecne dlouhemu prochazeni

        If TypeName(cand) = "MailItem" Then
            If cand.EntryID <> newEntryId Then
                Dim cs As String, cu As String, ct As Date
                cs = "": cu = "": ct = 0
                On Error Resume Next
                cs = cand.SenderEmailAddress
                cu = cand.subject
                ct = cand.ReceivedTime
                On Error GoTo ErrHandler

                Dim diffMin As Long
                diffMin = Abs(DateDiff("n", ct, t))

                ' Serazeno sestupne podle casu - jakmile jsme za oknem
                ' a jsme uz starsi nez novy mail, dal nema smysl hledat.
                If ct < t And diffMin > AUTO_DUPLICATE_WINDOW_MINUTES Then Exit For

                If diffMin <= AUTO_DUPLICATE_WINDOW_MINUTES Then
                    If LCase(Trim(cs)) = LCase(Trim(s)) And LCase(Trim(cu)) = LCase(Trim(u)) Then
                        matchFound = True
                        Exit For
                    End If
                End If
            End If
        End If
    Next cand

    If VERBOSE_LOGGING Then
        LogLine "CHECK | " & sn & " | " & u & " | " & Format(t, "YYYY-MM-DD HH:NN:SS") & _
                " | scanned=" & scanned & " | matchFound=" & matchFound
    End If

    If matchFound Then
        ArchiveAndDelete newItem, s, sn, u, t
    End If

    Exit Sub

ErrHandler:
    LogLine "CHYBA v CheckAndRemoveDuplicateSilent: " & Err.Number & " - " & Err.Description
End Sub

' Ulozi mail jako .msg do archivni slozky, zaloguje a smaze z Inboxu.
Private Sub ArchiveAndDelete(itm As Object, sender As String, senderName As String, subj As String, recTime As Date)
    EnsureFolderExists ARCHIVE_FOLDER_PATH

    Dim baseName As String
    baseName = Format(recTime, "YYYYMMDD-HHNNSS") & " - " & _
               SanitizeFileNamePart(senderName) & " - " & SanitizeFileNamePart(subj) & ".msg"

    Dim fullPath As String
    fullPath = GetUniqueFilePath(ArchiveFolder() & baseName)

    On Error Resume Next
    itm.SaveAs fullPath, olMSG
    If Err.Number <> 0 Then
        LogLine "CHYBA pri archivaci do .msg (" & Err.Number & " - " & Err.Description & "): " & fullPath
        Err.Clear
    End If
    On Error GoTo 0

    LogDuplicateAction senderName, subj, recTime
    itm.Delete
End Sub

Private Function ArchiveFolder() As String
    Dim p As String
    p = ARCHIVE_FOLDER_PATH
    If Right(p, 1) <> "\" Then p = p & "\"
    ArchiveFolder = p
End Function

Private Sub EnsureFolderExists(path As String)
    Dim p As String
    p = path
    If Right(p, 1) = "\" Then p = Left(p, Len(p) - 1)
    On Error Resume Next
    If Dir(p, vbDirectory) = "" Then MkDir p
    On Error GoTo 0
End Sub

' Odstrani znaky nepovolene v nazvech souboru a orizne prilis dlouhy text.
Private Function SanitizeFileNamePart(s As String) As String
    Dim badChars As String
    badChars = "\/:*?""<>|" & vbCrLf & vbTab
    Dim result As String
    result = Trim(s)

    Dim i As Long
    For i = 1 To Len(badChars)
        result = Replace(result, Mid(badChars, i, 1), "_")
    Next i

    If Len(result) = 0 Then result = "(neznamy)"
    If Len(result) > 80 Then result = Left(result, 80)

    SanitizeFileNamePart = result
End Function

' Pokud soubor uz existuje (dva duplikaty smazane ve stejnou vterinu),
' pripoji poradove cislo, at se navzajem neprepisou.
Private Function GetUniqueFilePath(basePath As String) As String
    If Dir(basePath) = "" Then
        GetUniqueFilePath = basePath
        Exit Function
    End If

    Dim ext As String, withoutExt As String
    ext = ".msg"
    withoutExt = Left(basePath, Len(basePath) - Len(ext))

    Dim n As Long
    n = 2
    Do While Dir(withoutExt & " (" & n & ")" & ext) <> ""
        n = n + 1
    Loop

    GetUniqueFilePath = withoutExt & " (" & n & ")" & ext
End Function

' Smaze z archivni slozky soubory starsi nez ARCHIVE_RETENTION_DAYS dni.
' Volano automaticky z InitDuplicateWatcher pri startu Outlooku,
' da se ale spustit i rucne (Alt+F8).
Public Sub PurgeOldArchivedEmails()
    Dim folderPath As String
    folderPath = ArchiveFolder()

    If Dir(folderPath, vbDirectory) = "" Then Exit Sub  ' archiv jeste neexistuje

    Dim cutoff As Date
    cutoff = Now - ARCHIVE_RETENTION_DAYS

    Dim fileName As String
    fileName = Dir(folderPath & "*.msg")

    Dim deletedCount As Long
    deletedCount = 0

    Do While fileName <> ""
        Dim fullPath As String
        fullPath = folderPath & fileName

        If FileDateTime(fullPath) < cutoff Then
            On Error Resume Next
            Kill fullPath
            If Err.Number = 0 Then
                deletedCount = deletedCount + 1
            Else
                LogLine "CHYBA pri mazani archivu: " & fullPath & " - " & Err.Description
                Err.Clear
            End If
            On Error GoTo 0
        End If

        fileName = Dir()
    Loop

    If deletedCount > 0 Then
        LogLine "PURGE ARCHIVU: smazano " & deletedCount & " souboru starsich nez " & ARCHIVE_RETENTION_DAYS & " dni."
    End If
End Sub

Private Function GetPrimaryInbox() As Outlook.MAPIFolder
    Dim store As Outlook.store
    For Each store In Application.Session.Stores
        If store.ExchangeStoreType = olPrimaryExchangeMailbox Then
            Set GetPrimaryInbox = store.GetDefaultFolder(olFolderInbox)
            Exit Function
        End If
    Next store
    Set GetPrimaryInbox = Nothing
End Function

Private Sub LogDuplicateAction(senderName As String, subj As String, recTime As Date)
    LogLine "DEL | " & senderName & " | " & subj & " | " & _
            Format(recTime, "YYYY-MM-DD HH:NN:SS")
End Sub

' Verejne rozhrani pro zapis vysledku rucniho cisteni z modulu EmailCleaner.
' Samotny zapis do souboru zustava centralizovany v tomto modulu.
Public Sub LogManualDuplicateAction(ByVal description As String)
    LogLine "MANUAL DEL | " & description
End Sub

Public Sub LogManualCleanupSummary(ByVal total As Long, ByVal found As Long, _
                                   ByVal deleted As Long, ByVal cancelled As Boolean)
    Dim status As String
    If cancelled Then
        status = "CANCELLED"
    Else
        status = "DONE"
    End If

    LogLine "MANUAL " & status & " | total=" & total & _
            " | found=" & found & " | deleted=" & deleted
End Sub

' Volano z ThisOutlookSession.InitDuplicateWatcher pri (re)startu watcheru.
Public Sub LogWatcherStart(inboxPath As String)
    LogLine "START | " & inboxPath & " | " & Format(Now, "YYYY-MM-DD HH:NN:SS")
End Sub

Private Sub LogLine(txt As String)
    Dim f As Integer
    f = FreeFile
    On Error Resume Next
    Open LOG_FILE_PATH For Append As #f
    Print #f, Format(Now, "YYYY-MM-DD HH:NN:SS") & "  " & txt
    Close #f
    On Error GoTo 0
End Sub

' Otevre log soubor v Poznamkovem bloku. Spustitelne primo z Alt+F8.
Public Sub OpenDuplicateLog()
    If Dir(LOG_FILE_PATH) = "" Then
        MsgBox "Log soubor jeste neexistuje (zatim se nic nelogovalo):" & vbCrLf & LOG_FILE_PATH, vbInformation
        Exit Sub
    End If
    Shell "notepad.exe """ & LOG_FILE_PATH & """", vbNormalFocus
End Sub
