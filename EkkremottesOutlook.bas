' ============================================================
' EkkremottesOutlook.bas
' Outlook VBA Macro — Άνοιγμα εκκρεμότητας από email
'
' ΕΓΚΑΤΑΣΤΑΣΗ:
' 1. Outlook → Alt+F11 → Project1 → Insert → Module
' 2. Επικόλληση αυτού του κώδικα
' 3. Αποθήκευση και επανεκκίνηση Outlook
' 4. Το "Νέα Εκκρεμότητα" εμφανίζεται στο right-click menu email
' ============================================================

Option Explicit

Private Const APP_URL As String = "https://atheneosinsurance.github.io/ekkremottes/"
Private Const FIREFOX_PATH As String = "C:\Program Files\Mozilla Firefox\firefox.exe"

' ── Κύρια συνάρτηση ─────────────────────────────────────────
Sub NewTicketFromEmail()
    Dim oMail As Outlook.MailItem
    Dim oSel  As Outlook.Selection

    ' Παίρνουμε το επιλεγμένο email
    Set oSel = Application.ActiveExplorer.Selection
    If oSel.Count = 0 Then
        MsgBox "Παρακαλώ επιλέξτε ένα email πρώτα.", vbExclamation, "Εκκρεμότητες"
        Exit Sub
    End If

    If Not TypeOf oSel.Item(1) Is Outlook.MailItem Then
        MsgBox "Το επιλεγμένο στοιχείο δεν είναι email.", vbExclamation, "Εκκρεμότητες"
        Exit Sub
    End If

    Set oMail = oSel.Item(1)

    ' ── Στοιχεία email ───────────────────────────────────────
    Dim sFrom    As String
    Dim sSubject As String
    Dim sDate    As String
    Dim sEntryID As String
    Dim sMsgID   As String

    sFrom    = GetSenderName(oMail)
    sSubject = oMail.Subject
    sDate    = Format(oMail.ReceivedTime, "yyyy-mm-dd")
    sEntryID = oMail.EntryID

    ' InternetMessageId — μοναδικό ID για Graph API search
    On Error Resume Next
    sMsgID = oMail.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x1035001F")
    On Error GoTo 0
    sMsgID = Replace(Replace(sMsgID, "<", ""), ">", "")

    ' ── Δημιουργία outlook: link ─────────────────────────────
    Dim sOutlookLink As String
    sOutlookLink = "outlook:entryid/" & sEntryID

    ' ── Κατασκευή URL ────────────────────────────────────────
    Dim sURL As String
    sURL = APP_URL & "?" & _
           "from=" & EncodeURL(sFrom) & _
           "&subject=" & EncodeURL(sSubject) & _
           "&date=" & EncodeURL(sDate) & _
           "&emaillink=" & EncodeURL(sOutlookLink) & _
           "&msgid=" & EncodeURL(sMsgID)

    ' ── Άνοιγμα Firefox ──────────────────────────────────────
    Dim sFirefox As String
    sFirefox = FIREFOX_PATH

    If Dir(sFirefox) = "" Then
        sFirefox = "C:\Program Files (x86)\Mozilla Firefox\firefox.exe"
    End If

    If Dir(sFirefox) = "" Then
        Shell "cmd /c start firefox """ & sURL & """", vbHide
    Else
        Shell """" & sFirefox & """ """ & sURL & """", vbNormalFocus
    End If

    Debug.Print "Ticket URL: " & sURL

End Sub

' ── Βοηθητικές συναρτήσεις ───────────────────────────────────

Private Function GetSenderName(oMail As Outlook.MailItem) As String
    On Error Resume Next
    Dim sName As String
    sName = oMail.SenderName
    If Len(Trim(sName)) = 0 Then
        sName = oMail.SenderEmailAddress
    End If
    GetSenderName = Trim(sName)
    On Error GoTo 0
End Function

Private Function EncodeURL(s As String) As String
    Dim i   As Integer
    Dim c   As String
    Dim out As String
    out = ""
    For i = 1 To Len(s)
        c = Mid(s, i, 1)
        Select Case c
            Case "A" To "Z", "a" To "z", "0" To "9", "-", "_", ".", "~"
                out = out & c
            Case " "
                out = out & "+"
            Case Else
                out = out & EncodeChar(c)
        End Select
    Next i
    EncodeURL = out
End Function

Private Function EncodeChar(c As String) As String
    Dim lCode As Long
    lCode = AscW(c)
    If lCode < 0 Then lCode = lCode + 65536

    If lCode < 128 Then
        EncodeChar = "%" & Right("0" & Hex(lCode), 2)
    ElseIf lCode < 2048 Then
        Dim b1 As Long, b2 As Long
        b1 = &HC0 Or (lCode \ 64)
        b2 = &H80 Or (lCode Mod 64)
        EncodeChar = "%" & Right("0" & Hex(b1), 2) & "%" & Right("0" & Hex(b2), 2)
    ElseIf lCode < 65536 Then
        Dim b3 As Long, b4 As Long, b5 As Long
        b3 = &HE0 Or (lCode \ 4096)
        b4 = &H80 Or ((lCode \ 64) Mod 64)
        b5 = &H80 Or (lCode Mod 64)
        EncodeChar = "%" & Right("0" & Hex(b3), 2) & "%" & Right("0" & Hex(b4), 2) & "%" & Right("0" & Hex(b5), 2)
    Else
        EncodeChar = ""
    End If
End Function

' ── Right-click menu setup ────────────────────────────────────
Public Sub SetupContextMenu()
    Dim oBar     As Office.CommandBar
    Dim oBtn     As Office.CommandBarButton
    Const MENU_TAG As String = "EkkremottesNewTicket"

    For Each oBar In Application.ActiveExplorer.CommandBars
        If oBar.Name = "Context Menu" Or oBar.Name = "Item Shortcut Menu" Then
            On Error Resume Next
            oBar.FindControl(Tag:=MENU_TAG).Delete
            On Error GoTo 0

            Set oBtn = oBar.Controls.Add(Type:=msoControlButton, Before:=1)
            With oBtn
                .Caption = "📋 Νέα Εκκρεμότητα"
                .Tag = MENU_TAG
                .Style = msoButtonIconAndCaption
                .OnAction = "NewTicketFromEmail"
                .BeginGroup = False
            End With
        End If
    Next oBar
End Sub

' ── Auto-setup στο άνοιγμα του Outlook ───────────────────────
' Βάλτε αυτές τις 2 γραμμές στο ThisOutlookSession module:
'
'   Private Sub Application_Startup()
'       Call SetupContextMenu
'   End Sub
