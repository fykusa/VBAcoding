' ============================================================
' Outlook VBA - Analyza velikosti schranky
' Pouziti: spust ShowMailboxReport nebo MarkLargeItems
' Vyzaduje UserForm "frmProgress" se temito objekty
'   label: lblStatus  - aktualni slozka
'   label: lblCounter - pocitadlo
'   text pole: txtResults - okno pro vypis vysledku
'   button:  - btnClose
' ============================================================
Option Explicit

Private Const TOP_ITEMS_COUNT As Long = 50
Private g_FolderCount As Long
Private g_ItemCount   As Long

Private Type FolderInfo
    Path        As String
    itemCount   As Long
    totalSize   As Double
End Type

' --- Pomocna procedura: aktualizuj progress okno ---
Private Sub UpdateProgress(line1 As String, line2 As String)
    frmProgress.lblStatus.Caption = line1
    frmProgress.lblCounter.Caption = line2
    DoEvents
End Sub

' --- Hlavni vstupni bod: prehled slozek ---
Public Sub ShowMailboxReport()
    Dim folders()   As FolderInfo
    Dim folderCount As Long
    folderCount = 0
    g_FolderCount = 0
    ReDim folders(0)

    frmProgress.txtResults.Visible = False
    frmProgress.btnClose.Visible = False
    frmProgress.Show vbModeless
    UpdateProgress "Probiha analyza slozek...", ""

    ' Pouze primarni osobni schranka (whitelist - vynecha SAP/sdilene inboxy)
    Dim store As Outlook.store
    For Each store In Application.Session.Stores
        If store.ExchangeStoreType = olPrimaryExchangeMailbox Then
            ScanFolder store.GetRootFolder, folders, folderCount
        End If
    Next store

    Unload frmProgress

    If folderCount = 0 Then
        MsgBox "Zadne slozky nenalezeny.", vbInformation
        Exit Sub
    End If

    ' Serad sestupne podle velikosti
    Dim i As Long, j As Long, tmp As FolderInfo
    For i = 0 To folderCount - 2
        For j = 0 To folderCount - 2 - i
            If folders(j).totalSize < folders(j + 1).totalSize Then
                tmp = folders(j): folders(j) = folders(j + 1): folders(j + 1) = tmp
            End If
        Next j
    Next i

    ShowFolderReport folders, folderCount
End Sub

' --- Rekurzivni prochazeni slozek ---
Private Sub ScanFolder(fldr As Outlook.MAPIFolder, _
                        ByRef arr() As FolderInfo, _
                        ByRef count As Long)
    g_FolderCount = g_FolderCount + 1
    UpdateProgress "Prochazim: " & fldr.name, _
                   "Slozky: " & g_FolderCount & "   |   Polozky: " & fldr.Items.count

    Dim totalSize As Double
    Dim itm As Object
    For Each itm In fldr.Items
        On Error Resume Next
        totalSize = totalSize + CDbl(itm.Size)
        On Error GoTo 0
    Next itm

    If count >= UBound(arr) + 1 Then ReDim Preserve arr(UBound(arr) + 50)
    arr(count).Path = fldr.FolderPath
    arr(count).itemCount = fldr.Items.count
    arr(count).totalSize = totalSize
    count = count + 1

    Dim sub_f As Outlook.MAPIFolder
    For Each sub_f In fldr.folders
        ScanFolder sub_f, arr, count
    Next sub_f
End Sub

' --- Zobrazeni vysledku v dialogu ---
Private Sub ShowFolderReport(arr() As FolderInfo, count As Long)
    Dim totalAll As Double
    Dim i        As Long
    For i = 0 To count - 1
        totalAll = totalAll + arr(i).totalSize
    Next i

    Dim msg As String
    msg = "PREHLED SLOZEK DLE VELIKOSTI" & vbCrLf
    msg = msg & "Celkem v schrance: " & FormatSize(totalAll) & vbCrLf
    msg = msg & String(62, "-") & vbCrLf
    msg = msg & PadRight("Velikost", 10) & "  " & PadRight("Polozky", 7) & "  Slozka" & vbCrLf
    msg = msg & String(62, "-") & vbCrLf

    Dim shown As Long
    For i = 0 To count - 1
        If arr(i).totalSize > 0 Then
            msg = msg & PadRight(FormatSize(arr(i).totalSize), 10) & "  " _
                      & PadRight(CStr(arr(i).itemCount), 7) & "  " _
                      & arr(i).Path & vbCrLf
            shown = shown + 1
            If shown >= 30 Then
                msg = msg & "  ... (dalsi prazdne slozky vynechany)" & vbCrLf
                Exit For
            End If
        End If
    Next i

    msg = msg & vbCrLf & "Pro oznaceni top " & TOP_ITEMS_COUNT & " nejvetsi mailu spust: MarkLargeItems"

    ' Zobraz vysledky v progress formulari (ne v modalnim MsgBox)
    With frmProgress
        .Caption = "Analyza schranky – vysledky"
        .lblStatus.Caption = "Hotovo. Celkem v schrance: " & FormatSize(totalAll)
        .lblCounter.Caption = "Zpracovano " & count & " slozek."
        .txtResults.Text = msg
        .txtResults.Visible = True
        .btnClose.Visible = True
        .Show vbModeless
    End With
End Sub

' --- Oznaceni top N nejvetsi mailu kategorii ---
Public Sub MarkLargeItems()
    Const CATEGORY As String = "! Velky mail"
    EnsureCategory CATEGORY, olCategoryColorRed

    Dim allItems() As Object
    Dim allSizes() As Double
    Dim itemCount  As Long
    itemCount = 0
    g_ItemCount = 0
    ReDim allItems(0)
    ReDim allSizes(0)

    frmProgress.txtResults.Visible = False
    frmProgress.btnClose.Visible = False
    frmProgress.Show vbModeless
    UpdateProgress "Hledam nejvetsi maily...", ""

    ' Pouze primarni osobni schranka (whitelist - vynecha SAP/sdilene inboxy)
    Dim store As Outlook.store
    For Each store In Application.Session.Stores
        If store.ExchangeStoreType = olPrimaryExchangeMailbox Then
            CollectItems store.GetRootFolder, allItems, allSizes, itemCount
        End If
    Next store

    Unload frmProgress

    If itemCount = 0 Then
        MsgBox "Zadne maily nenalezeny.", vbInformation
        Exit Sub
    End If

    ' Serad indexy sestupne podle velikosti
    Dim idx() As Long
    ReDim idx(itemCount - 1)
    Dim k As Long
    For k = 0 To itemCount - 1: idx(k) = k: Next k

    Dim ii As Long, jj As Long, tmpIdx As Long
    For ii = 0 To itemCount - 2
        For jj = 0 To itemCount - 2 - ii
            If allSizes(idx(jj)) < allSizes(idx(jj + 1)) Then
                tmpIdx = idx(jj): idx(jj) = idx(jj + 1): idx(jj + 1) = tmpIdx
            End If
        Next jj
    Next ii

    Dim marked As Long
    Dim limit  As Long
    If itemCount < TOP_ITEMS_COUNT Then limit = itemCount Else limit = TOP_ITEMS_COUNT

    For k = 0 To limit - 1
        Dim itm As Object
        Set itm = allItems(idx(k))
        On Error Resume Next
        itm.Categories = CATEGORY
        itm.Save
        On Error GoTo 0
        marked = marked + 1
    Next k

    Dim resultMsg As String
    resultMsg = "Oznaceno " & marked & " mailu kategorii """ & CATEGORY & """." & vbCrLf & _
                "Nejmensi oznaceny mail: " & FormatSize(allSizes(idx(limit - 1))) & vbCrLf & vbCrLf & _
                "Filtruj v Outlooku: Zobrazit > Seradit podle kategorie."

    With frmProgress
        .Caption = "Oznaceni dokonceno"
        .lblStatus.Caption = "Oznaceno " & marked & " mailu kategorii """ & CATEGORY & """."
        .lblCounter.Caption = "Nejmensi oznaceny mail: " & FormatSize(allSizes(idx(limit - 1)))
        .txtResults.Text = resultMsg
        .txtResults.Visible = True
        .btnClose.Visible = True
        .Show vbModeless
    End With
End Sub

' --- Rekurzivni sber mailu ---
Private Sub CollectItems(fldr As Outlook.MAPIFolder, _
                          ByRef arr() As Object, _
                          ByRef sizes() As Double, _
                          ByRef count As Long)
    g_ItemCount = g_ItemCount + fldr.Items.count
    UpdateProgress "Prochazim: " & fldr.name, _
                   "Maily nalezeny: " & g_ItemCount

    Dim itm As Object
    For Each itm In fldr.Items
        On Error Resume Next
        Dim sz As Double: sz = CDbl(itm.Size)
        On Error GoTo 0
        If sz > 0 Then
            If count >= UBound(arr) + 1 Then
                ReDim Preserve arr(UBound(arr) + 200)
                ReDim Preserve sizes(UBound(sizes) + 200)
            End If
            Set arr(count) = itm
            sizes(count) = sz
            count = count + 1
        End If
    Next itm

    Dim sub_f As Outlook.MAPIFolder
    For Each sub_f In fldr.folders
        CollectItems sub_f, arr, sizes, count
    Next sub_f
End Sub

' --- Vytvori kategorii pokud neexistuje ---
Private Sub EnsureCategory(catName As String, colorIdx As OlCategoryColor)
    Dim cat As Outlook.CATEGORY
    For Each cat In Application.Session.Categories
        If cat.name = catName Then Exit Sub
    Next cat
    Application.Session.Categories.Add catName, colorIdx
End Sub

Private Function FormatSize(bytes As Double) As String
    If bytes >= 1073741824 Then
        FormatSize = Format(bytes / 1073741824, "0.00") & " GB"
    ElseIf bytes >= 1048576 Then
        FormatSize = Format(bytes / 1048576, "0.0") & " MB"
    ElseIf bytes >= 1024 Then
        FormatSize = Format(bytes / 1024, "0") & " KB"
    Else
        FormatSize = CLng(bytes) & " B"
    End If
End Function

Private Function PadRight(s As String, width As Integer) As String
    PadRight = Left(s & Space(width), width)
End Function

