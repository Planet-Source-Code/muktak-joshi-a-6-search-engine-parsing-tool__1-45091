VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form frmMain 
   BackColor       =   &H00FFFFFF&
   Caption         =   "TechProtean MSearch 1.00"
   ClientHeight    =   6825
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8280
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6825
   ScaleWidth      =   8280
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin MSWinsockLib.Winsock sckClient 
      Index           =   0
      Left            =   3000
      Tag             =   "www.google.com"
      Top             =   240
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock sckClient 
      Index           =   1
      Left            =   3000
      Tag             =   "www.altavista.com"
      Top             =   720
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock sckClient 
      Index           =   2
      Left            =   3000
      Tag             =   "www.alltheweb.com"
      Top             =   1200
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock sckClient 
      Index           =   3
      Left            =   3000
      Tag             =   "web.ask.com"
      Top             =   1680
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock sckClient 
      Index           =   4
      Left            =   3000
      Tag             =   "s.teoma.com"
      Top             =   2160
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock sckClient 
      Index           =   5
      Left            =   3000
      Tag             =   "search.dmoz.org"
      Top             =   2640
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Frame fSearch 
      Appearance      =   0  'Flat
      BackColor       =   &H00404080&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6735
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2775
      Begin VB.CommandButton cmdStop 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         Caption         =   "Stop"
         Enabled         =   0   'False
         Height          =   375
         Left            =   1320
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   1200
         Width           =   615
      End
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   840
         Top             =   2520
      End
      Begin VB.CommandButton cmdGo 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         Caption         =   "Go"
         Default         =   -1  'True
         Height          =   375
         Left            =   2040
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   1200
         Width           =   615
      End
      Begin VB.TextBox txtSearch 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   600
         Width           =   2535
      End
      Begin VB.Label lblStandardLables 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Search the Web :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   210
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1380
      End
   End
   Begin SHDocVwCtl.WebBrowser wbResults 
      Height          =   6735
      Left            =   2760
      TabIndex        =   4
      Top             =   0
      Width           =   5535
      ExtentX         =   9763
      ExtentY         =   11880
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim SearchWhat() As String

Dim UseProxy As Boolean
Dim ProxyServer As String
Dim ProxyPort As String
'Dim TotalData() As String
Private Function TotalData(Index As Integer) As String

Dim FileNum
Dim sLine As String
Dim FinalData As String

FileNum = FreeFile
If Dir(App.Path + "\data" & Index & ".htm") <> "" Then
Open App.Path + "\data" & Index & ".htm" For Input As FileNum
While Not EOF(FileNum)
Line Input #FileNum, sLine
FinalData = FinalData & sLine
Wend
Close FileNum
End If
TotalData = FinalData
End Function
Private Sub cmdGo_Click()
Dim X As Integer
SearchWhat(0) = "http://www.google.com/search?q=" & txtSearch.Text 'http://www.google.com/search?q=
SearchWhat(1) = "http://www.altavista.com/web/results?q=" & txtSearch.Text
SearchWhat(2) = "http://www.alltheweb.com/search?&cat=web&cs=utf-8&_sb_lang=en&q=" & txtSearch.Text
SearchWhat(3) = "http://web.ask.com/web?q=" & txtSearch.Text
SearchWhat(4) = "http://s.teoma.com/search?q=" & txtSearch.Text
SearchWhat(5) = "http://search.dmoz.org/cgi-bin/search?search=" & txtSearch.Text
ResetPrevious
For X = 0 To 5
If Dir(App.Path + "\data" & X & ".htm") <> "" Then Kill App.Path + "\data" & X & ".htm"
If X <> 4 Then sckClient(X).Close
If UseProxy Then
If X <> 4 Then sckClient(X).Connect ProxyServer, ProxyPort
Else
If X <> 4 Then sckClient(X).Connect sckClient(X).Tag, 80
End If
Next
wbResults.Navigate2 "about:<table width=100% height=100%><tr><td valign=middle align=center><img src=" & Replace(App.Path, " ", "%20") & "\searching.gif></td></tr></table>"
txtSearch.Enabled = False
cmdStop.Enabled = True
cmdGo.Enabled = False
Timer1.Enabled = True
End Sub





Private Sub Command1_Click()

WriteHTML
End Sub

Private Sub cmdStop_Click()
Dim X As Integer
For X = 0 To 5
sckClient(X).Close
Next

txtSearch.Enabled = True
cmdGo.Enabled = True
cmdStop.Enabled = False
End Sub

Private Sub Form_Load()
    ReDim Preserve SearchWhat(6) As String
    'ReDim Preserve TotalData(6) As String
    ReDim Preserve FinalTitle(0) As String
    ReDim Preserve FinalURL(0) As String
    ReDim Preserve FinalDescription(0) As String
    
    ReDim Preserve Engine(0) As Integer
    ReDim Preserve EngineNames(0) As String
    wbResults.Navigate2 "about:blank"
End Sub

Private Sub Form_Resize()
'Me.Width = 8400
'Me.Height = 7230
If Me.WindowState <> vbMinimized Then
fSearch.Height = Me.ScaleHeight
wbResults.Move fSearch.Left + fSearch.Width, fSearch.Top, Me.ScaleWidth - fSearch.Width, Me.ScaleHeight
End If
End Sub

Private Sub sckClient_Close(Index As Integer)

sckClient(Index).Close
Select Case Index
        Case 0
            ParseGoogle TotalData(Index)
        Case 1
            ParseAltavista TotalData(Index)
        Case 2
            ParseAlltheWeb TotalData(Index)
        Case 3
            ParseJeeves TotalData(Index)
        Case 4
            ParseTeoma TotalData(Index)
        Case 5
            ParseDMOZ TotalData(Index)
End Select
WriteHTML
End Sub
Private Sub WriteData(Index As Integer, Data As String)
Dim FileNum
'Dim X As Integer
FileNum = FreeFile
Open App.Path + "\data" & Index & ".htm" For Append As FileNum
'For X = 0 To UBound(TotalData)
Print #FileNum, Data
'Next
Close FileNum
End Sub
Private Sub sckClient_Connect(Index As Integer)
    Dim strConnectString As String
If Index = 4 Then
sckClient(4).Close
Exit Sub
End If
    SearchWhat(Index) = Replace(SearchWhat(Index), " ", "+")
    strConnectString = "GET " & SearchWhat(Index) & " HTTP/1.0" & vbNewLine
    strConnectString = strConnectString & "Accept: image/gif, image/x-xbitmap, image/jpeg, image/pjpeg, */*" & vbNewLine
    strConnectString = strConnectString & "Accept-Language: en-us" & vbNewLine
    strConnectString = strConnectString & "Accept-Encoding: gzip, deflate" & vbNewLine
    strConnectString = strConnectString & "Host:" & sckClient(Index).Tag & vbNewLine
    strConnectString = strConnectString & "Connection: Keep-Alive" & vbNewLine & vbNewLine
    strConnectString = strConnectString & "User-Agent: Mozilla/4.0 (compatible; MSIE 5.5; Windows 98; Win 9x 4.90)"
    sckClient(Index).SendData strConnectString
    
End Sub

Private Sub sckClient_DataArrival(Index As Integer, ByVal bytesTotal As Long)
    Dim Data As String

    
    Debug.Print LenB(Data)
    sckClient(Index).GetData Data, vbString
    DoEvents
    
    'TotalData(Index) = TotalData(Index) & Data
    WriteData Index, Data

    End Sub

Private Function AllDisconnected() As Boolean
Dim X As Integer
Dim AllDis As Boolean
AllDis = True

For X = 0 To sckClient.UBound
If (sckClient(X).State = sckConnected) Or (sckClient(X).State = sckConnecting) Then

AllDis = False
End If
Next
AllDisconnected = AllDis
End Function

Private Sub Timer1_Timer()

If AllDisconnected Then
Timer1.Enabled = False
WriteHTML
txtSearch.Enabled = True
cmdGo.Enabled = True
cmdStop.Enabled = False
End If
End Sub

Private Sub txtSearch_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then cmdGo_Click
End Sub
