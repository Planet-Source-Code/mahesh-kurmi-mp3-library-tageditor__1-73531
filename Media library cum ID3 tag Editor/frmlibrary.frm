VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmLibrary 
   BackColor       =   &H80000005&
   Caption         =   "NeoPlayer Media Library"
   ClientHeight    =   9630
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   12180
   Icon            =   "frmlibrary.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9630
   ScaleWidth      =   12180
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ImageList ImgIconos 
      Left            =   900
      Top             =   2370
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   14
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmlibrary.frx":058A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmlibrary.frx":0F9C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmlibrary.frx":1336
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmlibrary.frx":1D48
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmlibrary.frx":3F2C
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmlibrary.frx":42C6
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmlibrary.frx":4860
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmlibrary.frx":4DFA
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmlibrary.frx":5194
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmlibrary.frx":572E
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmlibrary.frx":5CC8
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmlibrary.frx":6262
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmlibrary.frx":67FC
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmlibrary.frx":6B96
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   1
      Top             =   9315
      Width           =   12180
      _ExtentX        =   21484
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   18389
            Picture         =   "frmlibrary.frx":6F30
            Text            =   "Neo Player: Media Library"
            TextSave        =   "Neo Player: Media Library"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Picture         =   "frmlibrary.frx":74CA
            Text            =   "Records: "
            TextSave        =   "Records: "
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox PicClientArea 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   5085
      Left            =   0
      ScaleHeight     =   337
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   665
      TabIndex        =   2
      Top             =   0
      Width           =   10005
      Begin VB.CommandButton CmdSearch 
         Caption         =   "Command1"
         Height          =   315
         Left            =   8460
         TabIndex        =   6
         Top             =   90
         Width           =   1395
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   120
         TabIndex        =   3
         Text            =   "dil"
         Top             =   90
         Width           =   8160
      End
      Begin MSComctlLib.TreeView TreeFiles 
         Height          =   4245
         Left            =   120
         TabIndex        =   4
         Top             =   480
         Width           =   2685
         _ExtentX        =   4736
         _ExtentY        =   7488
         _Version        =   393217
         HideSelection   =   0   'False
         Indentation     =   176
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
         FullRowSelect   =   -1  'True
         ImageList       =   "ImgIconos"
         BorderStyle     =   1
         Appearance      =   1
         MousePointer    =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   4245
         Left            =   2850
         TabIndex        =   5
         Top             =   480
         Width           =   7035
         _ExtentX        =   12409
         _ExtentY        =   7488
         View            =   3
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   0   'False
         HideSelection   =   0   'False
         OLEDragMode     =   1
         OLEDropMode     =   1
         AllowReorder    =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ColHdrIcons     =   "ImageList1"
         ForeColor       =   16384
         BackColor       =   16777215
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         OLEDragMode     =   1
         OLEDropMode     =   1
         NumItems        =   11
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Filename"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Title"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Artist"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Album"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Genre"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Text            =   "Track No."
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   6
            Text            =   "Tracks Total"
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   7
            Text            =   "Year"
            Object.Width           =   1058
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   8
            Text            =   "Duration"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   9
            Text            =   "Bit Rate"
            Object.Width           =   2064
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Text            =   "Comments"
            Object.Width           =   4410
         EndProperty
      End
   End
   Begin MSComctlLib.ImageList Buttons 
      Left            =   1620
      Top             =   2100
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   14
      ImageHeight     =   19
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmlibrary.frx":7A64
            Key             =   "add"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmlibrary.frx":7ADB
            Key             =   "addi"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmlibrary.frx":7B59
            Key             =   "del"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmlibrary.frx":7C4A
            Key             =   "deli"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmlibrary.frx":7D3B
            Key             =   "next"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmlibrary.frx":7E2A
            Key             =   "prev"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmlibrary.frx":7F1A
            Key             =   "previ"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   1020
      Top             =   2700
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   8
      ImageHeight     =   7
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmlibrary.frx":801A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmlibrary.frx":80EC
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgLibrary 
      Left            =   1920
      Top             =   1680
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16777215
      _Version        =   393216
   End
   Begin VB.Label Label18 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Left            =   7740
      TabIndex        =   0
      Top             =   7620
      Width           =   7275
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuNewplaylist 
         Caption         =   "&New Playlist"
         Index           =   0
      End
      Begin VB.Menu mnuImportCurrentlist 
         Caption         =   "Import Current Playlist"
      End
      Begin VB.Menu mnuImportFile 
         Caption         =   "Import Playlist from file"
      End
      Begin VB.Menu mnubar8 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSavelist 
         Caption         =   "Save Playlist as"
      End
      Begin VB.Menu mnuExportList 
         Caption         =   "Export Playlist as.."
      End
      Begin VB.Menu mnubar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAddmedia 
         Caption         =   "&Add media to library.."
         Index           =   1
      End
      Begin VB.Menu mnubar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuview 
      Caption         =   "View"
      Begin VB.Menu mnurefresh 
         Caption         =   "Refresh"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuFavTracks 
         Caption         =   "Favourite Tracks"
      End
      Begin VB.Menu mnuHistory 
         Caption         =   "History"
      End
      Begin VB.Menu mnubar 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRemovemissing 
         Caption         =   "Remove missing files from Library"
      End
      Begin VB.Menu mnuEraseLibrary 
         Caption         =   "Remove All Entries from Library"
      End
      Begin VB.Menu mnubar9 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPreferences 
         Caption         =   "Preferences"
      End
   End
   Begin VB.Menu mnuhelp 
      Caption         =   "Help"
      Begin VB.Menu mnuplayerhelp 
         Caption         =   "Player help"
      End
      Begin VB.Menu mnubar4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "About"
      End
   End
   Begin VB.Menu mnuPopupListview 
      Caption         =   "PopupListview"
      Begin VB.Menu mnuPlay 
         Caption         =   "Play Selection as new Playlist"
      End
      Begin VB.Menu mnuEnqueue 
         Caption         =   "Enqueue Selection"
      End
      Begin VB.Menu mnuSend 
         Caption         =   "Send to"
         Begin VB.Menu mnuLstname 
            Caption         =   ""
            Index           =   0
         End
      End
      Begin VB.Menu mnubar7 
         Caption         =   "-"
      End
      Begin VB.Menu mnuremove 
         Caption         =   "Remove Item(s)"
      End
      Begin VB.Menu mnubar6 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExplore 
         Caption         =   "Explore item(s) Folder"
      End
      Begin VB.Menu mnuviewtag 
         Caption         =   "View/Edit Tag info.."
      End
   End
End
Attribute VB_Name = "frmLibrary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
'DefLng A-Z

Private Declare Function ShellExecute Lib "shell32" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function OpenClipboard Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function EmptyClipboard Lib "user32" () As Long
Private Declare Function CopyImage Lib "user32" (ByVal Handle As Long, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function SetClipboardData Lib "user32" (ByVal wFormat As Long, ByVal hMem As Long) As Long
Private Declare Function CloseClipboard Lib "user32" () As Long

Private Const SW_SHOW As Long = 5
Private Const CF_BITMAP As Long = 2
Private Const IMAGE_BITMAP As Long = 0
Private Const LR_COPYRETURNORG As Long = &H4

Private Const S_OTHER As String = "Other"

Private Const FILTER_BMP As String = "*.bmp;*.dib"
Private Const FILTER_GIF As String = "*.gif"
Private Const FILTER_JPEG As String = "*.jpeg;*.jpg;*.jpe;*.jfif;*.jfi;*.jif"
Private Const FILTER_PNG As String = "*.png"
Private Const FILTER_SUPPORTED As String = FILTER_BMP & ";" & FILTER_GIF & ";" & FILTER_JPEG & ";" & FILTER_PNG

Private Const MNU_COPY As Long = 0
Private Const MNU_PASTE As Long = 1

Private Const PASTE_TXT_1 As String = "&Paste"
Private Const PASTE_TXT_2 As String = PASTE_TXT_1 & " (this will change the current image)"

Dim myWindowState As Integer
Dim bInitialized As Boolean


Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Const LVM_FIRST = &H1000
Dim AutoSize As Boolean
Dim strSQL As String
Dim sConnectionString As String
Dim SQL As New ADODB.Connection
Dim rs As New ADODB.Recordset

Dim cnnMusic As ADODB.Connection
Dim cmd   As ADODB.Command
Dim bCurrentListImported As Boolean

' Name: StopFlicker
' Description:Avoid the Flickering
'Use this routine to stop a control (like a list or treeview) from flickering when it is getting it's data.
' By: Strider Solutions
'
'This code is copyrighted and has' limited warranties.Please see http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=4292&lngWId=1'for details.'**************************************

'Get more great source code from
' http://www.stridersolutions.com/products/cs/

Private Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long

Private Sub StopFlicker(ByVal lhWnd As Long)
Dim lRet As Long
'Object will not flicker - just be blank
lRet = LockWindowUpdate(lhWnd)
 End Sub
Private Sub Release()
Dim lRet As Long
lRet = LockWindowUpdate(0)
End Sub
Public Function BindToSQL(strSQL As String, Optional iSQL As ADODB.Connection, Optional iRs As ADODB.Recordset)
   On Error GoTo err_handler:
   'On Error Resume Next
    If iSQL Is Nothing Then
        Set iSQL = SQL
        If iSQL.state > 0 Then iSQL.Close
        Dim connstring As String
        connstring = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\Library\music.mdb;Persist Security Info=False"
        iSQL.Open connstring
    End If
    If strSQL = "" Then Exit Function
    
    If iRs Is Nothing Then Set iRs = rs
    If iRs.state > 0 Then iRs.Close
    iRs.Open strSQL, iSQL, adOpenKeyset, adLockOptimistic
    
    Dim sF As Field
    With ListView1
    .ListItems.Clear
    .ColumnHeaders.Clear
    For Each sF In iRs.Fields
        .ColumnHeaders.Add , , sF.name
    Next
    Dim x As Integer, Item
    Do While Not iRs.EOF
        For x = 0 To iRs.Fields.Count - 1
            If x = 0 Then
              If iRs(x).Value <> "" Then Set Item = ListView1.ListItems.Add(, , iRs(x).Value)
            Else
              If iRs(x).Value <> "" Then Item.SubItems(x) = iRs(x).Value
            End If
        Next
    iRs.MoveNext
    Loop
    .ColumnHeaders(1).Width = 1.9 * .ColumnHeaders(1).Width 'title
    .ColumnHeaders(2).Width = 1.2 * .ColumnHeaders(2).Width 'Artist
    .ColumnHeaders(5).Width = 0.5 * .ColumnHeaders(5).Width 'year
    .ColumnHeaders(6).Width = 0.6 * .ColumnHeaders(6).Width 'Length
    .ColumnHeaders(7).Width = 0.8 * .ColumnHeaders(7).Width 'playcount
    .ColumnHeaders(9).Width = 2 * .ColumnHeaders(9).Width 'year
    End With

Exit Function
err_handler:
    MsgBox "Error binding ListView to SQL" & vbNewLine & vbNewLine & "Error code:" & Err.Number & vbNewLine & "Error desc:" & Err.Description, vbCritical
End Function

Private Sub Command5_Click()
'On Error GoTo HELL
Dim scampos As String
Dim ArchivoINI As String
Dim stipo As String
Dim sSQl As String

        If Len(Trim(Text1)) = 0 Then Exit Sub
        scampos = "TITLE,ARTIST,ALBUM,GENRE,YEAR,LENGTH,PLAYCOUNT,PLAYEDLAST,FILE"
        sSQl = Replace(Text1, "'", "''", , , vbTextCompare)
        sSQl = "SELECT " & scampos & " FROM MUSIC WHERE TITLE LIKE '%" & sSQl & "%' OR ARTIST LIKE '%" & sSQl & "%' ORDER BY TITLE"
        BindToSQL sSQl
End Sub


Private Sub cmdSearch_Click()

Dim scampos As String
Dim sSQl As String
Dim sWhere As String
On Error GoTo err_handler:
If Len(Trim(Text1)) = 0 Then Exit Sub
   scampos = "TITLE,ARTIST,ALBUM,GENRE,YEAR,LENGTH,PLAYCOUNT,PLAYEDLAST,FILE"
   sSQl = Replace(Text1, "'", "''", , , vbTextCompare)
   sSQl = "SELECT " & scampos & " FROM MUSIC WHERE TITLE LIKE '%" & sSQl & "%' OR ARTIST LIKE '%" & sSQl & "%' ORDER BY TITLE"
   BindToSQL sSQl
   rs.Close
   sSQl = Replace(Text1, "'", "''", , , vbTextCompare)
   sWhere = "WHERE TITLE LIKE '%" & sSQl & "%' OR ARTIST LIKE '%" & sSQl & "%'" '' ORDER BY TITLE"

   'rs.Open "SELECT SUM(BYTES) AS TOTAL, SUM(SECONDS) AS TIEMPO " & " FROM MUSIC WHERE TITLE LIKE '%" & sSQL '& "%' OR ARTIST LIKE '%" & sSQL

   rs.Open "SELECT SUM(BYTES) AS TOTAL, SUM(SECONDS) AS TIEMPO FROM MUSIC " & sWhere, cnnMusic, adOpenForwardOnly, adLockReadOnly

   Dim lKilobytes As Long, lSeconds As Long
   lKilobytes = CLng(rs!Total / 1024)
   lSeconds = CLng(rs!TIEMPO)
   rs.Close
   
   '// UPDATE STATUS BAR
   Call UpdateStatusBar(lKilobytes, lSeconds)

Exit Sub
err_handler:
    MsgBox "Error binding ListView to SQL" & vbNewLine & vbNewLine & "Error code:" & Err.Number & vbNewLine & "Error desc:" & Err.Description, vbCritical

End Sub
Private Sub Form_Load()
    'On Error Resume Next
    Dim hBitmap As Long
   Dim i As Integer

  
    Dim t As Long
    Dim strT As String
    
    bInitialized = False
    
    XSize = GetSetting(Caption, "Window", "XSize", Width)
    If Err Then
        XSize = Width
        Err.Clear
    End If
    
    YSize = GetSetting(Caption, "Window", "YSize", Height)
    If Err Then
        YSize = Height
        Err.Clear
    End If
    
    XPos = GetSetting(Caption, "Window", "XPos", Left)
    If Err Then
        XPos = Left
        Err.Clear
    End If
    
    YPos = GetSetting(Caption, "Window", "YPos", Top)
    If Err Then
        YPos = Top
        Err.Clear
    End If
    
    myWindowState = GetSetting(Caption, "Window", "State", vbNormal)
    If Err Then
        myWindowState = vbNormal
        Err.Clear
    End If
    
    If XSize < 568 * Screen.TwipsPerPixelX Then XSize = 568 * Screen.TwipsPerPixelX
    If XSize > Screen.Width Then XSize = Screen.Width
    
    If YSize < 445 * Screen.TwipsPerPixelY Then YSize = 445 * Screen.TwipsPerPixelY
    If YSize > Screen.Height Then YSize = Screen.Height
    
    If XPos < 0 Then XPos = 0
    If XPos > Screen.Width - Width Then XPos = Screen.Width - Width
    
    If YPos < 0 Then YPos = 0
    If YPos > Screen.Height - Height Then YPos = Screen.Height - Height
    
    If myWindowState <> vbNormal And myWindowState <> vbMaximized Then _
        myWindowState = vbNormal
    
    If Width <> XSize Then _
        Width = XSize
    If Height <> YSize Then _
        Height = YSize
    
    If Left <> XPos Then _
        Left = XPos
    If Top <> YPos Then _
        Top = YPos
    
    If WindowState <> myWindowState Then _
        WindowState = myWindowState
    
    bInitialized = True
  mnuPopupListview.Visible = False
  ShowListViewColumnHeaderSortIcon ListView1
 
   
      
  Set cnnMusic = New ADODB.Connection

  With cnnMusic
    .Provider = "Microsoft.Jet.OLEDB.4.0"
    .Properties("Data Source") = App.Path & "Library\music.mdb"
    '.Properties("Jet OLEDB:Database Password") = "Licenciao159"
    .CursorLocation = adUseClient
    .Open
  End With
   Set cmd = New ADODB.Command
  cmd.ActiveConnection = cnnMusic
  LoadLibrary
 
  gHW = hWnd

    
End Sub

Private Sub Form_Resize()
    If WindowState <> vbMinimized Then
        If bInitialized Then myWindowState = WindowState
        PicClientArea.Width = ScaleWidth + 15
        PicClientArea.Height = ScaleHeight - 340
        Text1.Width = PicClientArea.ScaleWidth - 18 - CmdSearch.Width
        CmdSearch.Left = PicClientArea.ScaleWidth - CmdSearch.Width - 6
        ListView1.Width = PicClientArea.ScaleWidth - ListView1.Left - 6
        ListView1.Height = PicClientArea.ScaleHeight - 30 ' ScaleHeight - 800 '4 * ScaleHeight \ 5 - 2961
        TreeFiles.Height = ListView1.Height
    End If
End Sub

Private Sub Form_Terminate()
 Dim i As Long
    
    SaveSetting Caption, "Window", "XSize", XSize
    SaveSetting Caption, "Window", "YSize", YSize
    SaveSetting Caption, "Window", "XPos", XPos
    SaveSetting Caption, "Window", "YPos", YPos
    SaveSetting Caption, "Window", "State", myWindowState
    
    With ListView1
        SaveSetting Caption, "Columns", "SortKey", .SortKey
        SaveSetting Caption, "Columns", "SortOrder", .SortOrder
        
        For i = 1 To .ColumnHeaders.Count
            SaveSetting Caption, "ColumnPos", Format$(.ColumnHeaders(i).Position, "00"), i
            SaveSetting Caption, "Columns", Format$(i, "00"), .ColumnHeaders(i).Width
        Next
   End With
    
   SaveSetting Caption, "MP3s", "Directory", Text1.Text
   cnnMusic.Close
   Set cnnMusic = Nothing
End Sub

Public Sub Form_Unload(Cancel As Integer)
   frmLibrary.Visible = False
  
End Sub


Private Sub ListView1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Dim i As Long
   Dim idx As Long
   On Error Resume Next
   SortLvwOnLong frmLibrary.ListView1, ColumnHeader.Index
   ShowListViewColumnHeaderSortIcon ListView1
   EnsureSelVisible ListView1
End Sub


Private Sub ListView1_OLEDragDrop(Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, Y As Single)
Dim FileName As String
Dim Extension As String
Dim bfilefound As Boolean
Dim tt, icnt, iPos As Integer
If Effect <> 7 Then Exit Sub
For icnt = 1 To Data.Files.Count
   If FileExists(Data.Files(icnt)) Then
            'This function will add the index of this file added to the listview in order to
            'create a sequence of playback in normal sequential mode or shuffle mode
            iPos = InStrRev(Data.Files(icnt), ".")
            Extension = Mid(Data.Files(icnt), iPos + 1, Len(Data.Files(icnt)) - iPos)
        If UCase(Extension) = "MP3" Or UCase(Extension) = "MP2" Or UCase(Extension) = "MP1" Or UCase(Extension) = "WAV" Then
             iPos = InStrRev(Data.Files(icnt), "\")
             FileName = (Mid(Data.Files(icnt), iPos + 1, Len(Data.Files(icnt)) - iPos - 4))
             Call AddTrack(Data.Files(icnt))
        Else
             'addFilesfromDir (Data.Files(icnt))
        End If
        
   End If
1:
Next icnt

End Sub

Private Sub ListView1_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Me.MousePointer = 1
ListView1.MousePointer = 1
End Sub

Private Sub ListView1_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
On Error GoTo hell:
If Button = vbRightButton Then
If ListView1.SelectedItem.Text <> "" Then
   mnuPopupListview.Enabled = False
Else
   mnuPopupListview.Enabled = True
End If
   
PopupMenu mnuPopupListview
   
End If
hell:
End Sub



Private Sub mnuAddmedia_Click(Index As Integer)
  boolSearchShow = True
  frmSearch.bAddtracktoPlaylist = False 'Default addition to playlist to removed since it is call from library
  frmSearch.Show
End Sub

Private Sub mnuEnqueue_Click()
On Error GoTo hell

Dim i As Integer
Dim x As Integer
Dim sFile As String
   
If ListView1.ListItems.Count < 1 Then Exit Sub

For i = 1 To ListView1.ListItems.Count
  If ListView1.ListItems.Item(i).Selected = True Then
     sFile = ListView1.ListItems(i).SubItems(8)
     frmPLST.Add_track_to_Playlist (sFile)
  End If
Next i
frmPLST.Update_Plst_Scrollbar

hell:
End Sub

Private Sub mnuEraseLibrary_Click()
Dim k
If k = 7 Then Exit Sub
k = MsgBox("Are you sure you want to Erase all data from library?", vbYesNo, "NeoPlayer Media Library")
'k=6=Yes,  k=7=no
If k = 6 Then
On Error GoTo hell
Dim rs As New ADODB.Recordset
rs.Open "SELECT FILE FROM MUSIC", cnnMusic, adOpenDynamic, adLockOptimistic

Do Until rs.EOF
   rs.Delete
   rs.MoveNext
Loop
rs.Update
rs.Close
Set rs = Nothing
End If
hell:
LoadLibrary

End Sub

Private Sub mnuExit_Click()
 Call Form_Unload(0)
End Sub

Private Sub mnuExplore_Click()
Dim strPathExplore As String
If ListView1.ListItems.Count < 1 Then Exit Sub
strPathExplore = ListView1.SelectedItem.SubItems(8)
strPathExplore = Left(strPathExplore, InStrRev(strPathExplore, "\"))
Shell "explorer.exe " & strPathExplore, vbMaximizedFocus
End Sub

Private Sub mnuImportCurrentlist_Click()
bCurrentListImported = True
End Sub



Private Sub mnuPlay_Click()
On Error GoTo hell

Dim i As Integer
Dim x As Integer
Dim sFile As String
   
If ListView1.ListItems.Count < 1 Then Exit Sub
frmPLST.ClearList
For i = 1 To ListView1.ListItems.Count
  If ListView1.ListItems.Item(i).Selected = True Then
     sFile = ListView1.ListItems(i).SubItems(8)
     frmPLST.Add_track_to_Playlist (sFile)
  End If
Next i
frmPLST.Update_Plst_Scrollbar
CurrentTrack_Index = 0
      sFileMainPlaying = frmPLST.cList.exItem(CurrentTrack_Index)
      frmMain.PlayerIsPlaying = "true"
      frmPLST.ReinitializeList
      frmMain.Play
hell:
End Sub

Private Sub mnurefresh_Click()
LoadLibrary
End Sub

Private Sub mnuremove_Click()
'On Error GoTo Hell:
If ListView1.ListItems.Count < 1 Then Exit Sub
Dim sFile, sSQl As String
Dim i As Integer
For i = 1 To ListView1.ListItems.Count
    If ListView1.ListItems.Item(i).Selected = True Then
      sFile = ListView1.ListItems(i).SubItems(8)
      sSQl = Replace(sFile, "'", "''", , , vbTextCompare)
      cmd.CommandText = "DELETE FROM MUSIC WHERE FILE='" & sSQl & "'"
      cmd.Execute
    End If
Next i
  
TreeFiles_Click
hell:
End Sub

Private Sub mnuRemovemissing_Click()
On Error Resume Next
Dim rs As New ADODB.Recordset
rs.Open "SELECT FILE FROM MUSIC", cnnMusic, adOpenDynamic, adLockOptimistic

Do Until rs.EOF
   If Dir(rs!File) = "" Then
      rs.Delete
   End If
   rs.MoveNext
Loop
rs.Update
rs.Close
Set rs = Nothing
hell:
End Sub


Private Sub mnuviewtag_Click()
 'On Error Resume Next
     boolTagsShow = True
     frmPopUp.mnuTagEditor.Checked = boolTagsShow
       frmTags.Show
       DoEvents
       'frmTags.fileTags.Clear
       frmTags.vkFiletags.Clear
       frmTags.listRef.ListItems.Clear
       'don't let  vkfiletags draw itself repeatedly which it does on adding an item to it
       frmTags.vkFiletags.UnRefreshControl = True
       Dim i
       For i = 1 To ListView1.ListItems.Count
         If ListView1.ListItems.Item(i).Selected Then frmTags.Load_Tags ListView1.ListItems.Item(i).SubItems(8)  'add path to filetags
       Next
       
       If frmTags.vkFiletags.ListCount = 0 Then Exit Sub
       'Show tags of first item
       frmTags.Show_tags (1)
       'Show first item as selected
       frmTags.vkFiletags.Selected(1) = True 'index starts from 1 in vkListbox instead of zero
       'Now we can draw vkfiletags listbox
       frmTags.vkFiletags.UnRefreshControl = False
       frmTags.vkFiletags.Refresh
End Sub


Private Sub PicClientArea_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
 Me.MousePointer = vbDefault
 
 If Button = vbLeftButton And PicClientArea.MousePointer = 9 Then
  If x > 100 + TreeFiles.Left And x < TreeFiles.Left + PicClientArea.ScaleWidth - 100 Then
   TreeFiles.Width = x - TreeFiles.Left - 2
  ' ListView1.Left = X + 1
 '  ListView1.Width = PicClientArea.ScaleWidth - ListView1.Left - 6
   ListView1.Move x + 1, ListView1.Top, PicClientArea.ScaleWidth - ListView1.Left - 6, ListView1.Height
   'Form_Resize
  End If
End If
 If x > TreeFiles.Left + TreeFiles.Width And x < TreeFiles.Left + TreeFiles.Width + 6 Then
   PicClientArea.MousePointer = 9
 Else
   PicClientArea.MousePointer = vbDefault
 End If

End Sub

Private Sub PicClientArea_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
If PicClientArea.MousePointer = 9 Then Form_Resize
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Call cmdSearch_Click
End Sub

Private Sub TreeFiles_Click()
Dim Color As Single
Dim sAlbum As String
Dim sArtist As String
Dim sGenre As String
Dim sSQl As String
Dim stipo As String
Dim aEle() As String
Dim scampos As String
Dim sWhere As String
On Error Resume Next

'Stipo FOR SELECTED  KEY IN TREEVIEW
  stipo = Left(TreeFiles.SelectedItem.Key, 5)
    
  If stipo <> "kPlaE" And stipo <> "kRecI" And stipo <> "kRecP" And stipo <> "kTopH" And stipo <> "FL FO" And stipo <> "CDMCA" And stipo <> "kAll" And stipo <> "A  AL" And stipo <> "AA AR" And stipo <> "AA AL" And stipo <> "FL CA" And stipo <> "GAAGE" Then Exit Sub
    
  scampos = "TITLE,ARTIST,ALBUM,GENRE,YEAR,LENGTH,PLAYCOUNT,PLAYEDLAST,FILE"
    
  If stipo = "kAll" Then
    sWhere = "WHERE ONCD=FALSE"
    sSQl = "SELECT " & scampos & " FROM MUSIC " & sWhere
  End If
  
  '//CLICK ON ALBUMS
  If stipo = "A  AL" Then
     sAlbum = Right(TreeFiles.SelectedItem.Key, Len(TreeFiles.SelectedItem.Key) - 5)
     sWhere = "WHERE ONCD=FALSE AND ALBUM='" & sAlbum & "'"
     sSQl = "SELECT " & scampos & " FROM MUSIC " & sWhere
  End If
  
  '//CLICK ON ARTIST
  If stipo = "AA AR" Then
     sArtist = Right(TreeFiles.SelectedItem.Key, Len(TreeFiles.SelectedItem.Key) - 5)
     sWhere = "WHERE ONCD=FALSE AND ARTIST='" & sArtist & "'"
     sSQl = "SELECT " & scampos & " FROM MUSIC " & sWhere
  End If
    ' sArtist = aEle(1)
  
  '//CLICK ON ARTIST - ALBUM
  If stipo = "AA AL" Then
     aEle = Split(TreeFiles.SelectedItem.Key, "|", , vbTextCompare)
     sAlbum = aEle(2)
     sWhere = "WHERE ONCD=FALSE AND ARTIST='" & sArtist & "' AND ALBUM='" & sAlbum & "'"
     sSQl = "SELECT " & scampos & " FROM MUSIC " & sWhere
  End If
  
  '// CLICK ON FILE LOCATION
  If stipo = "FL CA" Or stipo = "FL FO" Then
     If TreeFiles.SelectedItem.Children = 0 Then
        sGenre = Right(TreeFiles.SelectedItem.Key, Len(TreeFiles.SelectedItem.Key) - 5)
        sGenre = Left(sGenre, Len(sGenre) - 1)
        sWhere = "WHERE ONCD=FALSE AND FILEPATH='" & sGenre & "'"
        sSQl = "SELECT " & scampos & " FROM MUSIC " & sWhere
     
     End If
  End If
  
  '// CD MEDIA
  If stipo = "CDMCA" Then
     If TreeFiles.SelectedItem.Children = 0 Then
        sGenre = Right(TreeFiles.SelectedItem.Key, Len(TreeFiles.SelectedItem.Key) - 5)
        sGenre = Left(sGenre, Len(sGenre) - 1)
        sWhere = "WHERE ONCD=TRUE AND FILEPATH='" & sGenre & "'"
        sSQl = "SELECT " & scampos & " FROM MUSIC " & sWhere
     Else
        If Len(TreeFiles.SelectedItem.Key) = 8 Then
           sGenre = Right(TreeFiles.SelectedItem.Key, 3)
           sWhere = "WHERE ONCD=TRUE AND DRIVE='" & sGenre & "'"
           sSQl = "SELECT " & scampos & " FROM MUSIC " & sWhere
        End If
     End If
  End If

  
   '//CLICK ON GENRES
  If stipo = "GAAGE" Then
     sGenre = Right(TreeFiles.SelectedItem.Key, Len(TreeFiles.SelectedItem.Key) - 5)
     sWhere = "WHERE ONCD=FALSE AND GENRE='" & sGenre & "'"
     sSQl = "SELECT " & scampos & " FROM MUSIC " & sWhere
  End If
  
  '//CLICK ON TOP HITS
  If stipo = "kTopH" Then
    sWhere = "WHERE ONCD=FALSE AND PLAYCOUNT>0 ORDER BY PLAYCOUNT DESC "
      'sSQL = "SELECT TOP 20 PLAYCOUNT," & sCampos & " FROM MUSIC " & sWhere
    sSQl = "SELECT " & scampos & " FROM MUSIC " & sWhere
   End If
  
     
  '//CLICK ON RECENTLY PLAYED
  If stipo = "kRecP" Then
      sWhere = "WHERE ONCD=FALSE AND PLAYEDLAST IS NOT NULL " & "ORDER BY PLAYEDLAST DESC"
      'sSQL = "SELECT TOP 10 PLAYEDLAST, " & sCampos & " FROM MUSIC " & sWhere
      sSQl = "SELECT " & scampos & " FROM MUSIC " & sWhere
  End If
  '*****iMPORTANT*****SPACE in sql SATAEMENET IS VERY IMP. "SELECT " IS CORRECT BUT "SELECT"IS WRONG
  '*****"SELECT TOP 20 LASTUPDATE, " IS CORRECT "SELECT TOP 20 LASTUPDATE," WRONG
    
  '//CLICK ON RECENTLY ADDED
  If stipo = "kRecI" Then
      sWhere = "WHERE ONCD=FALSE AND LASTUPDATE IS NOT NULL " & "ORDER BY LASTUPDATE DESC"
      'sSQL = "SELECT TOP 20 LASTUPDATE, " & sCampos & " FROM MUSIC " & sWhere 'CORRECT STATEMENT TO GET TOP 20 ENTRIES
      ' sSQL = "SELECT TOP 20 LASTUPDATE, " & sCampos & " FROM MUSIC " & sWhere 'CORRECT STATEMENT TO GET TOP 20 ENTRIES
      scampos = "TITLE,ARTIST,ALBUM,GENRE,YEAR,LENGTH,PLAYCOUNT,LASTUPDATE,FILE"
      sSQl = "SELECT " & scampos & " FROM MUSIC " & sWhere
  End If
  
  '//CLICK EN PLAYLISTS
  If stipo = "kPlaE" Then
      sGenre = Right(TreeFiles.SelectedItem.Key, Len(TreeFiles.SelectedItem.Key) - 5)
      'Cargar_PlayListTracks sGenre
      Exit Sub
  End If

 BindToSQL sSQl
     
 rs.Close
 If stipo = "kTopH" Then sWhere = "WHERE ONCD=FALSE AND PLAYCOUNT>0 "
 If stipo = "kRecP" Then sWhere = "WHERE ONCD=FALSE AND PLAYEDLAST IS NOT NULL "
 
 On Error Resume Next
 rs.Open "SELECT SUM(BYTES) AS TOTAL, SUM(SECONDS) AS TIEMPO FROM MUSIC " & sWhere, cnnMusic, adOpenForwardOnly, adLockReadOnly
   
    '// UPDATE STATUS BAR
 Dim lKilobytes As Long, lSeconds As Long
 lKilobytes = CLng(rs!Total / 1024)
 lSeconds = CLng(rs!TIEMPO)
 rs.Close
 Call UpdateStatusBar(lKilobytes, lSeconds)


End Sub

Public Sub LoadLibrary()
 Dim sClave As String
 Dim sLastNode As String
 Dim sLAlbum As String
 Dim sLArtist As String
 Dim sNode As String
 Dim sAlbum As String
 Dim sArtist As String
 Dim sGenre As String
 Dim rsFiles As New ADODB.Recordset
 Dim stipo As String
 On Error GoTo hell
  TreeFiles.Nodes.Clear
  ListView1.ListItems.Clear
  TreeFiles.Nodes.Add , , "kAll", "Local Audio", 3
  TreeFiles.Nodes.Add , , "kMediaLibrary", "Media Library", 5
  TreeFiles.Nodes.Add "kMediaLibrary", tvwChild, "kAlbum", "Album", 6
  TreeFiles.Nodes.Add "kMediaLibrary", tvwChild, "kArtist\Album", "Artist\Album", 7
  TreeFiles.Nodes.Add "kMediaLibrary", tvwChild, "kCDMedia", "CD Media", 8
  TreeFiles.Nodes.Add "kMediaLibrary", tvwChild, "kPath", "File Location", 9
  TreeFiles.Nodes.Add "kPath", tvwChild, "kFullPath", "Full Path", 9
  TreeFiles.Nodes.Add "kPath", tvwChild, "kFolder", "by Folder", 2
  TreeFiles.Nodes.Add "kMediaLibrary", tvwChild, "kGenre", "Genre", 13
  TreeFiles.Nodes.Add , , "kPlayList", "Play List", 14
  TreeFiles.Nodes.Add "kPlayList", tvwChild, "kTopHits", "Top Hits", 12
  TreeFiles.Nodes.Add "kPlayList", tvwChild, "kRecP", "Recently Played", 11
  TreeFiles.Nodes.Add "kPlayList", tvwChild, "kRecI", "Recently Imported", 10
  TreeFiles.Nodes.Add "kPlayList", tvwChild, "kPlaL", "Play Lists", 14
  
  
 cmd.CommandText = "SELECT DISTINCT GENRE,ARTIST FROM MUSIC WHERE ONCD =FALSE ORDER BY GENRE"
 
 Set rsFiles = cmd.Execute
  
 Dim rsArtist As New ADODB.Recordset
 Dim rsAlbum As New ADODB.Recordset

''  // GENRE
    Do Until rsFiles.EOF
       sGenre = CStr(Trim(rsFiles!Genre))
       If sGenre = "" Then sGenre = "Desconocido"
       
       If "GAAGE" & LCase(sGenre) <> sLastNode Then
           sLastNode = "GAAGE" & LCase(sGenre)
           TreeFiles.Nodes.Add "kGenre", tvwChild, sLastNode, sGenre, 13
       End If
       rsFiles.MoveNext
    Loop
  
  
 cmd.CommandText = "SELECT DISTINCT ALBUM FROM MUSIC WHERE ONCD=FALSE ORDER BY ALBUM"
 Set rsFiles = cmd.Execute
 sArtist = ""
 sAlbum = ""
  
''  // ALBUM
    Do Until rsFiles.EOF
       sAlbum = CStr(Trim(rsFiles!Album))
       If sAlbum = "" Then sAlbum = "Desconocido"
       
       If "A  AL" & LCase(sAlbum) <> sLastNode Then
           sLastNode = "A  AL" & LCase(sAlbum)
           TreeFiles.Nodes.Add "kAlbum", tvwChild, sLastNode, sAlbum, 6
       End If
       rsFiles.MoveNext
    Loop

'  // ARTIST - ALBUMS

 
 cmd.CommandText = "SELECT DISTINCT ARTIST FROM MUSIC WHERE ONCD=FALSE ORDER BY ARTIST"
 Set rsFiles = cmd.Execute
   sLastNode = ""
   sLAlbum = ""
     Do Until rsFiles.EOF
       sArtist = CStr(Trim(rsFiles!Artist))
       If sArtist = "" Then sArtist = "Desconocido"
       If "AA AR" & LCase(sArtist) <> sLastNode Then
           sLastNode = "AA AR" & LCase(sArtist)
           TreeFiles.Nodes.Add "kArtist\Album", tvwChild, sLastNode, sArtist, 7
           cmd.CommandText = "SELECT DISTINCT ALBUM FROM MUSIC WHERE ONCD=FALSE AND ARTIST='" & rsFiles!Artist & "'"
           Set rsAlbum = cmd.Execute
           If rsAlbum.RecordCount > 1 Then
           Do Until rsAlbum.EOF
                sAlbum = CStr(Trim(rsAlbum!Album))
                If sAlbum = "" Then sAlbum = "Desconocido"
                If "AA AL|" & LCase(sArtist) & "|" & LCase(sAlbum) <> sLAlbum Then
                sLAlbum = "AA AL|" & LCase(sArtist) & "|" & LCase(sAlbum)
                TreeFiles.Nodes.Add sLastNode, tvwChild, sLAlbum, sAlbum, 6
                End If
                rsAlbum.MoveNext
           
           Loop
           End If
           rsAlbum.Close
       
       End If
       rsFiles.MoveNext
    Loop
    
 '// FILE LOCATION
 Dim sKey As String, s As String
 Dim sPath() As String

 On Error Resume Next

 cmd.CommandText = "SELECT DISTINCT FILEPATH FROM MUSIC WHERE ONCD=FALSE"
 Set rsFiles = cmd.Execute
 sLastNode = ""
 sArtist = ""
 sAlbum = ""

 '// add albums folders
 Do Until rsFiles.EOF
    s = rsFiles!FilePath
        
    sPath = Split(s, "\", , vbTextCompare)
    TreeFiles.Nodes.Add "kFolder", tvwChild, "FL FO" & CStr(s & "\"), sPath(UBound(sPath)), 2
        
    If sLastNode <> sPath(0) Then
       TreeFiles.Nodes.Add "kFullPath", tvwChild, "FL CA" & CStr(sPath(0) & "\"), sPath(0), 1
       sLastNode = sPath(0)
    End If
    
    sKey = "FL CA" & sPath(0) & "\"
    Dim i As Integer
    For i = 1 To UBound(sPath)
       'If TreeFiles.Nodes(sKey).Children = 0 Then
          TreeFiles.Nodes.Add sKey, tvwChild, sKey & sPath(i) & "\", sPath(i), 2
        'End If
      If i = UBound(sPath) Then
        sKey = sKey & sPath(i)
      Else
        sKey = sKey & sPath(i) & "\"
      End If
    Next i
    rsFiles.MoveNext
    sKey = ""
 Loop
 
 
 '// CD MEDIA
 
 cmd.CommandText = "SELECT DISTINCT FILEPATH FROM MUSIC WHERE ONCD=TRUE"
 Set rsFiles = cmd.Execute
 sLastNode = ""
 sArtist = ""
 sAlbum = ""

 '// add albums folders
 Do Until rsFiles.EOF
    s = rsFiles!FilePath
        
    sPath = Split(s, "\", , vbTextCompare)
    
    If sLastNode <> sPath(0) Then
       TreeFiles.Nodes.Add "kCDMedia", tvwChild, "CDMCA" & CStr(sPath(0) & "\"), sPath(0), 1
       sLastNode = sPath(0)
    End If
    
    sKey = "CDMCA" & sPath(0) & "\"
    
    For i = 1 To UBound(sPath)
       'If TreeFiles.Nodes(sKey).Children = 0 Then
          TreeFiles.Nodes.Add sKey, tvwChild, sKey & sPath(i) & "\", sPath(i), 2
        'End If
      If i = UBound(sPath) Then
        sKey = sKey & sPath(i)
      Else
        sKey = sKey & sPath(i) & "\"
      End If
    Next i
    rsFiles.MoveNext
    sKey = ""
 Loop
 
 '-----------------------------------------------------------------------------------
'// buskar los archivos de playlist y agragarlos
'fPlayList.Pattern = "*.pls"


'If Dir(tAppConfig.AppConfig & "Library\", vbDirectory) <> "" Then
  'fPlayList.Path = tAppConfig.AppConfig & "Library\"
  
 'For i = 0 To fPlayList.ListCount - 1
   '   TreeFiles.Nodes.Add "kPlaL", tvwChild, "kPlaE" & Left(fPlayList.list(i), Len(fPlayList.list(i)) - 4), Left(fPlayList.list(i), Len(fPlayList.list(i)) - 4), 14
' Next i
'End If

 
' '// AGREGAR CD ROMS Y OTROS
'    Dim FS As New FileSystemObject
'    Dim dDrive As Drive
'    Dim dDrives As Drives
'
'
'    Set dDrives = FS.Drives
'
'    For Each dDrive In dDrives
'       'If dDrive.IsReady = True Then
'          Select Case dDrive.DriveType
'
'             Case 0 '/* Desconocido
'             Case 1 '/* Separable
'             Case 2 '/* Fijo
''                cboDrives.AddItem dDrive.DriveLetter & ": [" & dDrive.VolumeName & "]"
'             Case 3 '/* Red
'             Case 4 '/* CDROM
'                 If dDrive.IsReady = True Then
'                     sGenre = dDrive.DriveLetter & ": [" & dDrive.VolumeName & "]"
'                 Else
'                    sGenre = dDrive.DriveLetter
'                 End If
'
'                 TreeFiles.Nodes.Add "kCDS", tvwChild, "CDSFI" & dDrive.DriveLetter, sGenre, 1
'
'             Case 5 '/* Disco RAM
'          End Select
'      ' End If
'    Next
'
' Set FS = Nothing
TreeFiles.Nodes("kMediaLibrary").Expanded = True
TreeFiles.Nodes("kPlayList").Expanded = True

rsFiles.Close
rsArtist.Close
rsAlbum.Close
Set rsFiles = Nothing
Set rsArtist = Nothing
Set rsAlbum = Nothing

Dim sSQl As String
sSQl = "SELECT " & "TITLE,ARTIST,ALBUM,GENRE,YEAR,LENGTH,PLAYCOUNT,PLAYEDLAST,FILE" & " FROM MUSIC " & "WHERE ONCD=FALSE"
BindToSQL sSQl

rs.Close

Dim lKilobytes As Long, lSeconds As Long
rs.Open "SELECT SUM(BYTES) AS TOTAL, SUM(SECONDS) AS TIEMPO FROM MUSIC " & "WHERE ONCD=FALSE", cnnMusic, adOpenForwardOnly, adLockReadOnly
 
lKilobytes = CLng(rs!Total / 1024)
lSeconds = CLng(rs!TIEMPO)
rs.Close
Call UpdateStatusBar(lKilobytes, lSeconds)
Exit Sub
hell:

MsgBox Err.Description
End Sub

Public Sub UpdatePlaycount(sFile As String, Optional bcheckExistinLIBRARY As Boolean)
Dim rsAct As New ADODB.Recordset
Dim iContar As Integer
'On Error GoTo HELL
Dim s As String

If bcheckExistinLIBRARY = True Then Call AddTracktoLibrary(sFile, True, False) 'no need to show in list
s = Replace(sFile, "'", "''", , , vbTextCompare)
rsAct.Open "SELECT PLAYCOUNT,PLAYEDLAST FROM MUSIC WHERE FILE='" & s & "'", cnnMusic, adOpenDynamic, adLockPessimistic

If rsAct.RecordCount = 1 Then
    rsAct!PlayCount = rsAct!PlayCount + 1
    rsAct!PLayedLast = Now()
    rsAct.UpdateBatch adAffectCurrent
    iContar = rsAct!PlayCount + 1
    cmd.CommandText = "UPDATE MUSIC SET PLAYCOUNT=" & iContar & ",PLAYEDLAST='" & Now() & "' WHERE FILE='" & s & "'"
    cmd.Execute
End If
    rsAct.Close
    Set rsAct = Nothing
hell:
'MsgBox err.Description
End Sub
Public Function AddTracktoLibrary(sFile As String, Optional bcheckExistinLIBRARY As Boolean = False, Optional bshowinList As Boolean = False) As Boolean
  Dim cFile As New cMP3
  Dim rst As New ADODB.Recordset
  Dim s As String
  Dim lSeconds As Long
  Dim sTitle As String, sArtist As String, sAlbum As String, sGenre As String, sYear As String, sComment As String
  Dim bUpdateLib As Boolean
  bUpdateLib = True
  
  On Error Resume Next
  sFile = Replace(sFile, "'", "''", , , vbTextCompare)
  If bcheckExistinLIBRARY Then
    bUpdateLib = Not Exist_in_Library(sFile)
  Else
    bUpdateLib = True
  End If
  
  'no need to check for id3 info if not to be added eg. dblclick playlist event
  If bshowinList = False And bUpdateLib = False Then AddTracktoLibrary = False: Exit Function
  
  cFile.Read_MPEGInfo = True
  cFile.Read_File_Tags sFile
  sTitle = Replace(cFile.Title, "'", " ", , , vbTextCompare)
  sArtist = Replace(cFile.Artist, "'", " ", , , vbTextCompare)
  sAlbum = Replace(cFile.Album, "'", " ", , , vbTextCompare)
  sYear = Replace(cFile.Year, "'", " ", , , vbTextCompare)
  sGenre = Replace(cFile.Genre, "'", " ", , , vbTextCompare)
  sComment = Replace(cFile.Comment, "'", " ", , , vbTextCompare)
            
  If sTitle = "" Then sTitle = GetFileTitle(sFile)
  If sArtist = "" Then sArtist = "Unknown"
  If sAlbum = "" Then sAlbum = "Unknown"
  If sYear = "" Then sYear = Year(Now())
  If sGenre = "" Then sGenre = "Other"
  If sComment = "" Then sComment = "Uncommented"
            
  If bshowinList = True Then
    Dim Item
    Set Item = ListView1.ListItems.Add(, , sTitle)
    Item.SubItems(1) = sArtist
    Item.SubItems(2) = sAlbum
    Item.SubItems(3) = sGenre
    Item.SubItems(4) = sYear
    Item.SubItems(5) = cFile.MPEG_DurationTime
    Item.SubItems(6) = ""
    Item.SubItems(7) = ""
    Item.SubItems(8) = sFile
  End If

If bUpdateLib = True Then
  rst.Open "SELECT * FROM Music", cnnMusic, adOpenDynamic, adLockOptimistic
  rst.AddNew
  rst!File = sFile
  rst!Title = sTitle
  rst!Artist = sArtist
  rst!Album = sAlbum
  rst!Year = sYear
  rst!Genre = sGenre
  rst!Comments = sComment
  rst!length = cFile.MPEG_DurationTime
  rst!BYTES = cFile.FileSize
  rst!Seconds = cFile.DurationInSecs
' rst!LastUpdate = cFile.LastUpdate
  rst!PlayCount = 0
  rst!Quality = cFile.Quality
  rst!Situation = cFile.Situation
' rst!Mood = cFile.Mood
  rst!FilePath = sFile
  'rst!OnCD = bCDROM
  rst!Drive = Left(sFile, 3)
  rst.Update
  rst.Close
  AddTracktoLibrary = True 'Successfully added return TRUE
 End If

 AddTracktoLibrary = True
End Function


Public Sub DisplayTrack_in_List(sFile As String)
'Just display track with id3 info in listview without changing any data in library
  Dim cFile As New cMP3
  Dim lSeconds As Long
  Dim sTitle As String, sArtist As String, sAlbum As String, sGenre As String, sYear As String, sComment As String

  cFile.Read_MPEGInfo = True
        
  cFile.Read_File_Tags sFile
  sTitle = Replace(cFile.Title, "'", " ", , , vbTextCompare)
  sArtist = Replace(cFile.Artist, "'", " ", , , vbTextCompare)
  sAlbum = Replace(cFile.Album, "'", " ", , , vbTextCompare)
  sYear = Replace(cFile.Year, "'", " ", , , vbTextCompare)
  sGenre = Replace(cFile.Genre, "'", " ", , , vbTextCompare)
  sComment = Replace(cFile.Comment, "'", " ", , , vbTextCompare)
            
  If sTitle = "" Then sTitle = GetFileTitle(sFile)
  If sArtist = "" Then sArtist = "Unknown"
  If sAlbum = "" Then sAlbum = "Unknown"
  If sYear = "" Then sYear = Year(Now())
  If sGenre = "" Then sGenre = "Other"
  If sComment = "" Then sComment = "Uncommented"
            
  Dim Item
  Set Item = ListView1.ListItems.Add(, , sTitle)
  Item.SubItems(1) = sArtist
  Item.SubItems(2) = sAlbum
  Item.SubItems(3) = sGenre
  Item.SubItems(4) = sYear
  Item.SubItems(5) = cFile.MPEG_DurationTime
  Item.SubItems(6) = ""
  Item.SubItems(7) = ""
  Item.SubItems(8) = sFile
End Sub

Public Function Exist_in_Library(sFile As String) As Boolean
  Dim rst As New ADODB.Recordset
  If sFile = "" Then Exit Function
  sFile = Replace(sFile, "'", "''", , , vbTextCompare)
  rst.Open "SELECT PLAYCOUNT,PLAYEDLAST FROM MUSIC WHERE FILE='" & sFile & "'", cnnMusic, adOpenDynamic, adLockPessimistic
  'If Recodcount>0 then file is existing somewhere in library
  If rst.RecordCount <> 0 Then Exist_in_Library = True
  rst.Close
End Function

Public Sub UpdateStatusBar(lKilobytes As Long, lSeconds As Long)
    StatusBar1.Panels(2).Text = "RECORDS:[ " & ListView1.ListItems.Count & " ]   -   "
    StatusBar1.Panels(1).Text = "RECORD:[" + str(ListView1.SelectedItem.Index) + "] " + ListView1.SelectedItem

    'Dim lKILOBYTES As Long
    Dim DD As Long, HH As Long, MM As Long, ss As Long, sTempTime As String
    
    sTempTime = "SIZE: [ " & Format(lKilobytes, "000,000") & " KB. ]   -   "
    
    StatusBar1.Panels(2).Text = StatusBar1.Panels(2).Text + sTempTime
    sTempTime = ""
    DD = lSeconds \ 86400     ' Days
    lSeconds = Abs(lSeconds - (DD * 86400))
    HH = lSeconds \ 3600      ' Hours
    MM = lSeconds \ 60 Mod 60 ' Minutes
    ss = lSeconds Mod 60      ' Seconds
    sTempTime = "TIME:[ "
    If DD > 0 Then sTempTime = sTempTime & DD & " days. "
    If HH > 0 Then sTempTime = sTempTime & HH & " Hr. "
    StatusBar1.Panels(2).Text = StatusBar1.Panels(2).Text + sTempTime & MM & " Min. " & Format$(ss, "00") & " Sec. ]"

End Sub
