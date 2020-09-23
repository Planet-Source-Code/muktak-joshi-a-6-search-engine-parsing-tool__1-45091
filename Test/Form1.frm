VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5880
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7755
   LinkTopic       =   "Form1"
   ScaleHeight     =   5880
   ScaleWidth      =   7755
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   1335
   End
   Begin VB.ListBox List3 
      Height          =   1815
      Left            =   5160
      TabIndex        =   3
      Top             =   3960
      Width           =   2415
   End
   Begin VB.ListBox List2 
      Height          =   1815
      Left            =   2640
      TabIndex        =   2
      Top             =   3960
      Width           =   2415
   End
   Begin VB.ListBox List1 
      Height          =   1815
      Left            =   120
      TabIndex        =   1
      Top             =   3960
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      Height          =   3255
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "Form1.frx":0000
      Top             =   600
      Width           =   7455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Title() As String
Dim URL() As String
Dim Description() As String
Private Sub Command1_Click()
'Text1.Text = Replace(Text1.Text, vbCrLf,"", , , vbTextCompare)
'Text1.Text = Replace(Text1.Text, "  ", " ")
'Text1.Text = Replace(Text1.Text, "  ", " ")
'Text1.Text = Replace(Text1.Text, "  ", " ")
'Text1.Text = Replace(Text1.Text, "  ", " ")

    'ParseGoogle (Text1.Text)
    'ParseAltavista (Text1.Text)
    ParseAlltheWeb (Text1.Text)
    'ParseJeeves (Text1.Text)
    'ParseTeoma (Text1.Text)
    'ParseDMOZ (Text1.Text)
    
    
End Sub

Private Sub Form_Load()
ReDim URL(0) As String
ReDim Title(0) As String
ReDim Description(0) As String
End Sub



Private Sub WriteHTML()
'MsgBox UBound(Title)
Dim FileNum
Dim x As Integer
Dim sLine As String
FileNum = FreeFile
Open App.Path + "\tmp.htm" For Output As FileNum
For x = 0 To UBound(Title)
sLine = "<table width=100%>"
sLine = sLine & "<tr><td bgcolor=gray><a href=" & URL(x) & ">" & Title(x) & "</a></td></tr>"
sLine = sLine & "<tr><td>" & Description(x) & "</td></tr>"
Print #FileNum, sLine
Next
Close FileNum
'frmMain.wbResults.Navigate App.Path + "\tmp.htm"
End Sub






Public Sub ParseGoogle(Result As String)
    Dim ResultBlock() As String
    Dim pos As String
    Dim x As Integer

    ReDim Preserve ResultBlock(0) As String
    pos = 1
    While pos <> 0
        pos = InStr(1, Result, "<P class=g>", vbTextCompare)
        Result = Right(Result, Len(Result) - pos)
        ReDim Preserve ResultBlock(UBound(ResultBlock) + 1) As String
        ResultBlock(UBound(ResultBlock)) = GetLeftWord(Result, "<p class=g>", False)
        Result = Right(Result, Len(Result) - Len(ResultBlock(UBound(ResultBlock))))
    Wend
    For x = 1 To UBound(ResultBlock) - 1
        ReDim Preserve URL(UBound(URL) + 1) As String
        pos = InStr(1, ResultBlock(x), "href=", vbTextCompare)
        If pos <> 0 Then pos = pos + 4
        ResultBlock(x) = Right(ResultBlock(x), Len(ResultBlock(x)) - pos)
        'pos = InStr(1, ResultBlock(X), Chr(34), vbTextCompare)
        pos = InStr(1, ResultBlock(x), ">", vbTextCompare)
        URL(UBound(URL)) = Left$(ResultBlock(x), pos - 1)
        ResultBlock(x) = Right(ResultBlock(x), Len(ResultBlock(x)) - pos)

        ReDim Preserve Title(UBound(Title) + 1) As String
        Title(UBound(Title)) = GetLeftWord(ResultBlock(x), "</A>", False)
        ResultBlock(x) = Right(ResultBlock(x), Len(ResultBlock(x)) - Len(Title(UBound(Title))))
        Title(UBound(Title)) = Replace(Title(UBound(Title)), vbCrLf, "", , , vbTextCompare)
        Title(UBound(Title)) = Replace(Title(UBound(Title)), "<B>", "", , , vbTextCompare)
        Title(UBound(Title)) = Replace(Title(UBound(Title)), "</B>", "", , , vbTextCompare)
        Title(UBound(Title)) = Replace(Title(UBound(Title)), "<BR>", "", , , vbTextCompare)
        Title(UBound(Title)) = Replace(Title(UBound(Title)), "<FONT size=-1>", "", , , vbTextCompare)
        Title(UBound(Title)) = Replace(Title(UBound(Title)), "</FONT>", "", , , vbTextCompare)
        Title(UBound(Title)) = Replace(Title(UBound(Title)), ">", "", , , vbTextCompare)
        Title(UBound(Title)) = Replace(Title(UBound(Title)), "&amp;", "", , , vbTextCompare)



        ReDim Preserve Description(UBound(Description) + 1) As String
        pos = InStr(1, ResultBlock(x), "<FONT color=#008000", vbTextCompare)
        Description(UBound(Description)) = Left$(ResultBlock(x), pos)
        'If InStr(1, Description(UBound(Description)), "<SPAN") <> 0 Then
        Description(UBound(Description)) = GetLeftWord(ResultBlock(x), "<SPAN", True)
        Description(UBound(Description)) = Replace(Description(UBound(Description)), vbCrLf, "", , , vbTextCompare)
        Description(UBound(Description)) = Replace(Description(UBound(Description)), "<B>", "", , , vbTextCompare)
        Description(UBound(Description)) = Replace(Description(UBound(Description)), "</B>", "", , , vbTextCompare)
        Description(UBound(Description)) = Replace(Description(UBound(Description)), "<BR>", "", , , vbTextCompare)
        Description(UBound(Description)) = Replace(Description(UBound(Description)), "<FONT size=-1", "", , , vbTextCompare)
        Description(UBound(Description)) = Replace(Description(UBound(Description)), "</FONT>", "", , , vbTextCompare)

        Description(UBound(Description)) = Replace(Description(UBound(Description)), "</A>", "", , , vbTextCompare)
        Description(UBound(Description)) = Replace(Description(UBound(Description)), ">", "", , , vbTextCompare)
        Description(UBound(Description)) = Replace(Description(UBound(Description)), "&amp;", "", , , vbTextCompare)
        Description(UBound(Description)) = GetLeftWord(Description(UBound(Description)), "<FONT", False)

        If Trim(Description(UBound(Description))) = "" Then Description(UBound(Description)) = "No Description"

     
    Next
    Call WriteHTML
End Sub


Public Sub ParseAltavista(Result As String)
    Dim ResultBlock() As String
    Dim pos As String
    Dim x As Integer

    ReDim Preserve ResultBlock(0) As String
    pos = 1
    While pos <> 0
        pos = InStr(1, Result, "class=result", vbTextCompare)
        Result = Right(Result, Len(Result) - pos)
        ReDim Preserve ResultBlock(UBound(ResultBlock) + 1) As String
        ResultBlock(UBound(ResultBlock)) = GetLeftWord(Result, "</TD>", False)
        Result = Right(Result, Len(Result) - Len(ResultBlock(UBound(ResultBlock))))
    Wend
    For x = 1 To UBound(ResultBlock) - 1
        ReDim Preserve URL(UBound(URL) + 1) As String
        pos = InStr(1, ResultBlock(x), "status='", vbTextCompare)
        If pos <> 0 Then pos = pos + 7
        ResultBlock(x) = Right(ResultBlock(x), Len(ResultBlock(x)) - pos)
        pos = InStr(1, ResultBlock(x), "'", vbTextCompare)
        URL(UBound(URL)) = Left$(ResultBlock(x), pos - 1)
        ResultBlock(x) = Right(ResultBlock(x), Len(ResultBlock(x)) - pos)

        ReDim Preserve Title(UBound(Title) + 1) As String
        Title(UBound(Title)) = GetLeftWord(ResultBlock(x), "</A>", False)
        ResultBlock(x) = Right(ResultBlock(x), Len(ResultBlock(x)) - Len(Title(UBound(Title))))
        Title(UBound(Title)) = GetRightWord(Title(UBound(Title)), Chr(34) & ">", False)

        Title(UBound(Title)) = Replace(Title(UBound(Title)), vbCrLf, "", , , vbTextCompare)

        Title(UBound(Title)) = Replace(Title(UBound(Title)), "  ", " ")
        Title(UBound(Title)) = Replace(Title(UBound(Title)), "  ", " ")
        Title(UBound(Title)) = Replace(Title(UBound(Title)), "  ", " ")
        Title(UBound(Title)) = Replace(Title(UBound(Title)), "  ", " ")

        Title(UBound(Title)) = Replace(Title(UBound(Title)), ">", "", , , vbTextCompare)
        Title(UBound(Title)) = Replace(Title(UBound(Title)), "&amp;", "", , , vbTextCompare)
        If Left$(Title(UBound(Title)), 24) = "http://www.altavista.com" Then Title(UBound(Title)) = "No Title"


        ReDim Preserve Description(UBound(Description) + 1) As String
        ResultBlock(x) = Replace(ResultBlock(x), vbCrLf, "", , , vbTextCompare)
        ResultBlock(x) = Replace(ResultBlock(x), "  ", " ")
        ResultBlock(x) = Replace(ResultBlock(x), "  ", " ")
        ResultBlock(x) = Replace(ResultBlock(x), "  ", " ")


        pos = InStr(1, ResultBlock(x), "<SPAN class=s>", vbTextCompare)
        If pos <> 0 Then pos = pos + 13

        ResultBlock(x) = Right(ResultBlock(x), Len(ResultBlock(x)) - pos)
        Description(UBound(Description)) = GetLeftWord(ResultBlock(x), "<SPAN", False)
        Description(UBound(Description)) = Replace(Description(UBound(Description)), "<B>", "", , , vbTextCompare)
        Description(UBound(Description)) = Replace(Description(UBound(Description)), "</B>", "", , , vbTextCompare)
        Description(UBound(Description)) = Replace(Description(UBound(Description)), "<BR>", "", , , vbTextCompare)
        Description(UBound(Description)) = Replace(Description(UBound(Description)), "</A>", "", , , vbTextCompare)
        Description(UBound(Description)) = Replace(Description(UBound(Description)), ">", "", , , vbTextCompare)
        Description(UBound(Description)) = Replace(Description(UBound(Description)), "&amp;", "", , , vbTextCompare)


        If Trim(Description(UBound(Description))) = "" Then Description(UBound(Description)) = "No Description"

        
    Next
    Call WriteHTML
End Sub

Public Sub ParseJeeves(Result As String)
    Dim ResultBlock() As String
    Dim pos As String
    Dim x As Integer

    ReDim Preserve ResultBlock(0) As String
    pos = 1
    pos = InStr(1, Result, "search result", vbTextCompare)
    Result = Right(Result, Len(Result) - pos)

    While pos <> 0
        pos = InStr(1, Result, "smallerMargin", vbTextCompare)
        Result = Right(Result, Len(Result) - pos)
        ReDim Preserve ResultBlock(UBound(ResultBlock) + 1) As String
        ResultBlock(UBound(ResultBlock)) = GetLeftWord(Result, "smallerMargin", False)
        Result = Right(Result, Len(Result) - Len(ResultBlock(UBound(ResultBlock))))
    Wend
    For x = 1 To UBound(ResultBlock) - 1
        ReDim Preserve URL(UBound(URL) + 1) As String
        'Pos = InStr(1, ResultBlock(X), Chr(34))
        URL(UBound(URL)) = GetLeftWord(ResultBlock(x), "return ss('go to", False)
        ResultBlock(x) = Right(ResultBlock(x), Len(ResultBlock(x)) - Len(URL(UBound(URL))))
        URL(UBound(URL)) = GetLeftWord(ResultBlock(x), "')", False)
        ResultBlock(x) = Right(ResultBlock(x), Len(ResultBlock(x)) - Len(URL(UBound(URL))))
        URL(UBound(URL)) = GetRightWord(URL(UBound(URL)), "to ", False)
        URL(UBound(URL)) = Right(URL(UBound(URL)), Len(URL(UBound(URL))) - 2)
        ResultBlock(x) = Right(ResultBlock(x), Len(ResultBlock(x)) - pos)


        ReDim Preserve Title(UBound(Title) + 1) As String
        Title(UBound(Title)) = GetLeftWord(ResultBlock(x), "</A>", False)
        ResultBlock(x) = Right(ResultBlock(x), Len(ResultBlock(x)) - Len(Title(UBound(Title))))
        Title(UBound(Title)) = GetRightWord(Title(UBound(Title)), ">", False)

        Title(UBound(Title)) = Replace(Title(UBound(Title)), vbCrLf, "", , , vbTextCompare)
        Title(UBound(Title)) = Replace(Title(UBound(Title)), "  ", " ")
        Title(UBound(Title)) = Replace(Title(UBound(Title)), "  ", " ")
        Title(UBound(Title)) = Replace(Title(UBound(Title)), "  ", " ")
        Title(UBound(Title)) = Replace(Title(UBound(Title)), "  ", " ")
        Title(UBound(Title)) = Replace(Title(UBound(Title)), "&amp;", "", , , vbTextCompare)



        ReDim Preserve Description(UBound(Description) + 1) As String
        pos = InStr(1, ResultBlock(x), "<DIV", vbTextCompare)
        If pos <> 0 Then pos = pos + 4

        ResultBlock(x) = Right(ResultBlock(x), Len(ResultBlock(x)) - pos)

        Description(UBound(Description)) = GetLeftWord(ResultBlock(x), "</DIV>", False)



        Description(UBound(Description)) = Replace(Description(UBound(Description)), vbCrLf, "", , , vbTextCompare)
        Description(UBound(Description)) = Replace(Description(UBound(Description)), "<EM>", "", , , vbTextCompare)
        Description(UBound(Description)) = Replace(Description(UBound(Description)), "</EM>", "", , , vbTextCompare)
        Description(UBound(Description)) = Replace(Description(UBound(Description)), "<DIV>", "", , , vbTextCompare)
        Description(UBound(Description)) = Replace(Description(UBound(Description)), "</DIV>", "", , , vbTextCompare)
        Description(UBound(Description)) = Replace(Description(UBound(Description)), "  ", " ")
        Description(UBound(Description)) = Replace(Description(UBound(Description)), "  ", " ")
        Description(UBound(Description)) = Replace(Description(UBound(Description)), "  ", " ")
        Description(UBound(Description)) = Replace(Description(UBound(Description)), "  ", " ")

        Description(UBound(Description)) = Replace(Description(UBound(Description)), ">", "", , , vbTextCompare)
        Description(UBound(Description)) = Replace(Description(UBound(Description)), "&amp;", "", , , vbTextCompare)


        If Trim(Description(UBound(Description))) = "" Then Description(UBound(Description)) = "No Description"

        
    Next
    Call WriteHTML
End Sub
Public Sub ParseAlltheWeb(Result As String)
    Dim ResultBlock() As String
    Dim pos As String
    Dim x As Integer
    
    Result = Replace(Result, Chr(34), "")
        
    ReDim Preserve ResultBlock(0) As String
    pos = 1
    While pos <> 0
        pos = InStr(1, Result, "<p class=result", vbTextCompare)
        Result = Right(Result, Len(Result) - pos)
        ReDim Preserve ResultBlock(UBound(ResultBlock) + 1) As String
        ResultBlock(UBound(ResultBlock)) = GetLeftWord(Result, "<p class=result", False)
        Result = Right(Result, Len(Result) - Len(ResultBlock(UBound(ResultBlock))))
    Wend
    For x = 1 To UBound(ResultBlock) - 1
        ReDim Preserve URL(UBound(URL) + 1) As String
        pos = InStr(1, ResultBlock(x), "class=resURL", vbTextCompare)
        If pos <> 0 Then pos = pos + 11
        
        URL(UBound(URL)) = Right(ResultBlock(x), Len(ResultBlock(x)) - pos)
        'pos = InStr(1, ResultBlock(x), "'", vbTextCompare)
        'URL(UBound(URL)) = Left$(ResultBlock(x), pos - 1)
        URL(UBound(URL)) = GetLeftWord(URL(UBound(URL)), "</span>", False)
       URL(UBound(URL)) = Replace(URL(UBound(URL)), ">", "")
        'ResultBlock(x) = Right(ResultBlock(x), Len(ResultBlock(x)) - pos)

        ReDim Preserve Title(UBound(Title) + 1) As String
        Title(UBound(Title)) = GetLeftWord(ResultBlock(x), "</A>", False)
        ResultBlock(x) = Right(ResultBlock(x), Len(ResultBlock(x)) - Len(Title(UBound(Title))))
        Title(UBound(Title)) = GetRightWord(Title(UBound(Title)), ">", True)
        Title(UBound(Title)) = Replace(Title(UBound(Title)), vbCrLf, "", , , vbTextCompare)
        Title(UBound(Title)) = Replace(Title(UBound(Title)), "<SPAN class=hlight>", "", , , vbTextCompare)
        Title(UBound(Title)) = Replace(Title(UBound(Title)), "</SPAN>", "", , , vbTextCompare)
        Title(UBound(Title)) = Replace(Title(UBound(Title)), ">", "", , , vbTextCompare)



        ReDim Preserve Description(UBound(Description) + 1) As String

        Description(UBound(Description)) = GetLeftWord(ResultBlock(x), "<br", False)
Description(UBound(Description)) = GetRightWord(Description(UBound(Description)), "<br", False)
MsgBox Description(UBound(Description))
        Description(UBound(Description)) = Replace(Description(UBound(Description)), vbCrLf, "", , , vbTextCompare)
        Description(UBound(Description)) = Replace(Description(UBound(Description)), "<SPAN class=hlight>", "", , , vbTextCompare)
        Description(UBound(Description)) = Replace(Description(UBound(Description)), "<SPAN class=resDescr>", "", , , vbTextCompare)
        Description(UBound(Description)) = Replace(Description(UBound(Description)), "<SPAN class=resTeaser>", "", , , vbTextCompare)
        Description(UBound(Description)) = Replace(Description(UBound(Description)), "<SPAN class=resDescrLabel>", "", , , vbTextCompare)

        Description(UBound(Description)) = Replace(Description(UBound(Description)), "</SPAN>", "", , , vbTextCompare)

        Description(UBound(Description)) = Replace(Description(UBound(Description)), "<B>", "", , , vbTextCompare)
        Description(UBound(Description)) = Replace(Description(UBound(Description)), "</B>", "", , , vbTextCompare)
        Description(UBound(Description)) = Replace(Description(UBound(Description)), "<BR>", "", , , vbTextCompare)
        Description(UBound(Description)) = Replace(Description(UBound(Description)), "<FONT size=-1", "", , , vbTextCompare)
        Description(UBound(Description)) = Replace(Description(UBound(Description)), "</FONT>", "", , , vbTextCompare)

        Description(UBound(Description)) = Replace(Description(UBound(Description)), "</A>", "", , , vbTextCompare)
        Description(UBound(Description)) = Replace(Description(UBound(Description)), "&amp;", "", , , vbTextCompare)
        Description(UBound(Description)) = GetLeftWord(Description(UBound(Description)), "<FONT", False)
        Description(UBound(Description)) = Replace(Description(UBound(Description)), ">", "", , , vbTextCompare)
        'Description(UBound(Description)) = Right(Description(UBound(Description)), Len(Description(UBound(Description))) - 10)
        If Trim(Description(UBound(Description))) = "" Then Description(UBound(Description)) = "No Description"

    Next
    Call WriteHTML
End Sub

Public Sub ParseTeoma(Result As String)
    Dim ResultBlock() As String
    Dim pos As String
    Dim x As Integer

    ReDim Preserve ResultBlock(0) As String
    pos = 1
    While pos <> 0
        pos = InStr(1, Result, "<DIV><SPAN class=resultTxt>", vbTextCompare)
        Result = Right(Result, Len(Result) - pos)
        ReDim Preserve ResultBlock(UBound(ResultBlock) + 1) As String
        ResultBlock(UBound(ResultBlock)) = GetLeftWord(Result, "</DIV>", False)
        Result = Right(Result, Len(Result) - Len(ResultBlock(UBound(ResultBlock))))
    Wend
    For x = 1 To UBound(ResultBlock) - 1
        ReDim Preserve URL(UBound(URL) + 1) As String
        pos = InStr(1, ResultBlock(x), Chr(34) & ">", vbTextCompare)
        URL(UBound(URL)) = Left$(ResultBlock(x), pos - 1)
        ResultBlock(x) = Right(ResultBlock(x), Len(ResultBlock(x)) - pos)

        URL(UBound(URL)) = GetRightWord(URL(UBound(URL)), "u=", False)
        URL(UBound(URL)) = Right(URL(UBound(URL)), Len(URL(UBound(URL))) - 1)

        ReDim Preserve Title(UBound(Title) + 1) As String
        Title(UBound(Title)) = GetLeftWord(ResultBlock(x), "</A>", False)
        ResultBlock(x) = Right(ResultBlock(x), Len(ResultBlock(x)) - Len(Title(UBound(Title))))
        Title(UBound(Title)) = GetRightWord(Title(UBound(Title)), Chr(34), True)
        Title(UBound(Title)) = Replace(Title(UBound(Title)), vbCrLf, "", , , vbTextCompare)
        Title(UBound(Title)) = Replace(Title(UBound(Title)), "<SPAN class=bold", "", , , vbTextCompare)

        'Title(UBound(Title)) = Replace(Title(UBound(Title)), "<SPAN","", , , vbTextCompare)
        Title(UBound(Title)) = Replace(Title(UBound(Title)), "</SPAN>", "", , , vbTextCompare)
        Title(UBound(Title)) = Replace(Title(UBound(Title)), ">", "", , , vbTextCompare)



        ReDim Preserve Description(UBound(Description) + 1) As String

        Description(UBound(Description)) = GetLeftWord(ResultBlock(x), "<A", False)
        Description(UBound(Description)) = Replace(Description(UBound(Description)), vbCrLf, "", , , vbTextCompare)
        Description(UBound(Description)) = GetLeftWord(Description(UBound(Description)), "<SPAN class=baseURI>", False)
        Description(UBound(Description)) = Replace(Description(UBound(Description)), "<SPAN class=baseURI>", "", , , vbTextCompare)

        Description(UBound(Description)) = Replace(Description(UBound(Description)), "<SPAN class=bold>", "", , , vbTextCompare)


        Description(UBound(Description)) = Replace(Description(UBound(Description)), "</SPAN>", "", , , vbTextCompare)


        Description(UBound(Description)) = Replace(Description(UBound(Description)), "<BR>", "", , , vbTextCompare)


        Description(UBound(Description)) = Replace(Description(UBound(Description)), "</A>", "", , , vbTextCompare)
        Description(UBound(Description)) = Replace(Description(UBound(Description)), "&amp;", "", , , vbTextCompare)

        Description(UBound(Description)) = Replace(Description(UBound(Description)), ">", "", , , vbTextCompare)
        If InStr(1, Description(UBound(Description)), "[", vbTextCompare) <> 0 Then Description(UBound(Description)) = ""
        If Trim(Description(UBound(Description))) = "" Then Description(UBound(Description)) = "No Description"

       
    Next
    Call WriteHTML
End Sub
Public Sub ParseDMOZ(Result As String)
    Dim ResultBlock() As String
    Dim pos As String
    Dim x As Integer

    ReDim Preserve ResultBlock(0) As String
    pos = 1
    pos = InStr(1, Result, "Open Directory Sites", vbTextCompare)
    If pos <> 0 Then Result = Right(Result, Len(Result) - pos)
    While pos <> 0
        pos = InStr(1, Result, "<LI>", vbTextCompare)
        Result = Right(Result, Len(Result) - pos)
        ReDim Preserve ResultBlock(UBound(ResultBlock) + 1) As String
        ResultBlock(UBound(ResultBlock)) = GetLeftWord(Result, "<LI>", False)
        Result = Right(Result, Len(Result) - Len(ResultBlock(UBound(ResultBlock))))
    Wend
    For x = 1 To UBound(ResultBlock) - 1
        ReDim Preserve URL(UBound(URL) + 1) As String
        pos = InStr(1, ResultBlock(x), Chr(34), vbTextCompare)
        ResultBlock(x) = Right(ResultBlock(x), Len(ResultBlock(x)) - pos)
        pos = InStr(1, ResultBlock(x), Chr(34), vbTextCompare)
        URL(UBound(URL)) = Left$(ResultBlock(x), pos - 1)
        ResultBlock(x) = Right(ResultBlock(x), Len(ResultBlock(x)) - pos)


        ReDim Preserve Title(UBound(Title) + 1) As String
        Title(UBound(Title)) = GetLeftWord(ResultBlock(x), "</A>", False)
        ResultBlock(x) = Right(ResultBlock(x), Len(ResultBlock(x)) - Len(Title(UBound(Title))))

        Title(UBound(Title)) = Replace(Title(UBound(Title)), vbCrLf, "", , , vbTextCompare)


        Title(UBound(Title)) = Replace(Title(UBound(Title)), "<B>", "", , , vbTextCompare)
        Title(UBound(Title)) = Replace(Title(UBound(Title)), "</B>", "", , , vbTextCompare)
        Title(UBound(Title)) = Replace(Title(UBound(Title)), ">", "", , , vbTextCompare)
        Title(UBound(Title)) = Replace(Title(UBound(Title)), "  ", " ")
        Title(UBound(Title)) = Replace(Title(UBound(Title)), "  ", " ")
        Title(UBound(Title)) = Replace(Title(UBound(Title)), "  ", " ")



        ReDim Preserve Description(UBound(Description) + 1) As String

        Description(UBound(Description)) = GetLeftWord(ResultBlock(x), "<SMALL>", False)
        Description(UBound(Description)) = GetRightWord(Description(UBound(Description)), "</A>", False)
        Description(UBound(Description)) = GetRightWord(Description(UBound(Description)), "&nbsp; -", False)

        Description(UBound(Description)) = Replace(Description(UBound(Description)), vbCrLf, "", , , vbTextCompare)




        Description(UBound(Description)) = Replace(Description(UBound(Description)), "nbsp;", "", , , vbTextCompare)
        Description(UBound(Description)) = Replace(Description(UBound(Description)), "<BR>", "", , , vbTextCompare)
        Description(UBound(Description)) = Replace(Description(UBound(Description)), "/A>", "", , , vbTextCompare)
        Description(UBound(Description)) = Replace(Description(UBound(Description)), "</A>", "", , , vbTextCompare)
        Description(UBound(Description)) = Replace(Description(UBound(Description)), "&amp;", "", , , vbTextCompare)
        Description(UBound(Description)) = Replace(Description(UBound(Description)), "<B>", "", , , vbTextCompare)
        Description(UBound(Description)) = Replace(Description(UBound(Description)), "</B>", "", , , vbTextCompare)
        Description(UBound(Description)) = Replace(Description(UBound(Description)), "-", "", , , vbTextCompare)
        If InStr(1, Description(UBound(Description)), "<img", vbTextCompare) <> 0 Then Description(UBound(Description)) = GetRightWord(Description(UBound(Description)), ">", False)
        Description(UBound(Description)) = Replace(Description(UBound(Description)), ">", "", , , vbTextCompare)
        Description(UBound(Description)) = Replace(Description(UBound(Description)), "& ", "", , , vbTextCompare)
        Description(UBound(Description)) = Trim(Description(UBound(Description)))
        
        If Trim(Description(UBound(Description))) = "" Then Description(UBound(Description)) = "No Description"

      
    Next
    Call WriteHTML
End Sub

Private Sub Class_Initialize()

End Sub


