Attribute VB_Name = "modParseFunctions"
Option Explicit

Public FinalTitle() As String
Public FinalURL() As String
Public FinalDescription() As String
Public Engine() As Integer
Public EngineNames() As String
'Public Title() As String
'Public URL() As String
'Public Description() As String


Public Sub WriteHTML()

Dim FileNum
Dim X As Integer
Dim sLine As String

fnSort Engine, 0
DoEvents
FileNum = FreeFile
If Dir(App.Path + "\tmp.htm") <> "" Then Kill App.Path + "\tmp.htm"
Open App.Path + "\tmp.htm" For Output As FileNum
If UBound(FinalTitle) = 0 Then
sLine = "<font face=arial size=2><h3>Connection Error</h3><hr><br>Unable to Connect to Search Engines. Please Check your Connection and Try again.</font>"
Else
sLine = "<font face=arial><h3>Search Results</h3></font><hr><br>"
End If
Print #FileNum, sLine
sLine = ""


For X = 0 To UBound(FinalTitle) - 1
sLine = sLine & "<table width=100% cellpadding=3 cellspacing=3 border=0>"
sLine = sLine & "<tr><td bgcolor=lightyellow style='font-family:arial;font-size:11px;font-weight:bold;color:gray;border:1px solid;'> <a target=_blank href=" & FinalURL(X) & ">" & FinalTitle(X) & "</a><i> - From " & EngineNames(X) & "</i></td></tr>" '( " & Engine(X) & " )
sLine = sLine & "<tr><td bgcolor=#FFCC99 style='font-family:arial;font-size:11px;color:black;border:1px solid;'>" & FinalDescription(X) & "</td></tr>"
Print #FileNum, sLine
sLine = ""
Next

Close FileNum
frmMain.wbResults.Navigate App.Path + "\tmp.htm"

    
End Sub
Public Sub ResetPrevious()
ReDim FinalTitle(0) As String
ReDim FinalURL(0) As String
ReDim FinalDescription(0) As String
ReDim Engine(0) As Integer
ReDim EngineNames(0) As String
End Sub
Private Sub AddToFinal(URL As String, Title As String, Description As String, SEngine As String)
Dim X As Integer
Dim Index As Integer
Dim Found As Boolean
For X = 1 To UBound(FinalURL)
If (URL <> "") And (FinalURL(X) <> "") Then
'If (InStr(1, URL, FinalURL(X), vbTextCompare) <> 0) Then 'Or (InStr(1, URL, FinalURL(X), vbTextCompare) <> 0) Then
If (Trim(URL) = Trim(FinalURL(X))) Or (Trim(Title) = Trim(FinalTitle(X))) Then
Found = True
Index = X
End If
End If
'Exit For

Next

If Found Then
If InStr(1, EngineNames(Index), SEngine, vbTextCompare) = 0 Then
Engine(Index) = Val(Engine(Index)) + 1
EngineNames(Index) = EngineNames(Index) & ", " & SEngine
End If
Else
ReDim Preserve FinalTitle(UBound(FinalTitle) + 1) As String
ReDim Preserve FinalURL(UBound(FinalURL) + 1) As String
ReDim Preserve FinalDescription(UBound(FinalDescription) + 1) As String
ReDim Preserve Engine(UBound(Engine) + 1) As Integer
ReDim Preserve EngineNames(UBound(EngineNames) + 1) As String

FinalTitle(UBound(FinalTitle)) = Title
FinalURL(UBound(FinalURL)) = URL
FinalDescription(UBound(FinalDescription)) = Description
EngineNames(UBound(EngineNames)) = SEngine
Engine(UBound(Engine)) = 1
End If
End Sub
Public Sub ParseGoogle(Result As String)
    Dim ResultBlock() As String
    Dim pos As String
    Dim X As Integer
    
    Dim Title() As String
    Dim URL() As String
    Dim Description() As String
    
    ReDim Preserve Title(0) As String
    ReDim Preserve URL(0) As String
    ReDim Preserve Description(0) As String
    

    ReDim Preserve ResultBlock(0) As String
    pos = 1
    While pos <> 0
        pos = InStr(1, Result, "<P class=g>", vbTextCompare)
        Result = Right(Result, Len(Result) - pos)
        ReDim Preserve ResultBlock(UBound(ResultBlock) + 1) As String
        ResultBlock(UBound(ResultBlock)) = GetLeftWord(Result, "<p class=g>", False)
        Result = Right(Result, Len(Result) - Len(ResultBlock(UBound(ResultBlock))))
    Wend
    For X = 1 To UBound(ResultBlock) - 1
        ReDim Preserve URL(UBound(URL) + 1) As String
        pos = InStr(1, ResultBlock(X), "href=", vbTextCompare)
        If pos <> 0 Then pos = pos + 4
        ResultBlock(X) = Right(ResultBlock(X), Len(ResultBlock(X)) - pos)
        'pos = InStr(1, ResultBlock(X), Chr(34), vbTextCompare)
        pos = InStr(1, ResultBlock(X), ">", vbTextCompare)
        URL(UBound(URL)) = Left$(ResultBlock(X), pos - 1)
        ResultBlock(X) = Right(ResultBlock(X), Len(ResultBlock(X)) - pos)

        ReDim Preserve Title(UBound(Title) + 1) As String
        Title(UBound(Title)) = GetLeftWord(ResultBlock(X), "</A>", False)
        ResultBlock(X) = Right(ResultBlock(X), Len(ResultBlock(X)) - Len(Title(UBound(Title))))
        Title(UBound(Title)) = Replace(Title(UBound(Title)), vbCrLf, "", , , vbTextCompare)
        Title(UBound(Title)) = Replace(Title(UBound(Title)), "<B>", "", , , vbTextCompare)
        Title(UBound(Title)) = Replace(Title(UBound(Title)), "</B>", "", , , vbTextCompare)
        Title(UBound(Title)) = Replace(Title(UBound(Title)), "<BR>", "", , , vbTextCompare)
        Title(UBound(Title)) = Replace(Title(UBound(Title)), "<FONT size=-1>", "", , , vbTextCompare)
        Title(UBound(Title)) = Replace(Title(UBound(Title)), "</FONT>", "", , , vbTextCompare)
        Title(UBound(Title)) = Replace(Title(UBound(Title)), ">", "", , , vbTextCompare)
        Title(UBound(Title)) = Replace(Title(UBound(Title)), "&amp;", "", , , vbTextCompare)

Title(UBound(Title)) = Title(UBound(Title)) '& " - From Google"

        ReDim Preserve Description(UBound(Description) + 1) As String
        pos = InStr(1, ResultBlock(X), "<FONT color=#008000", vbTextCompare)
        Description(UBound(Description)) = Left$(ResultBlock(X), pos)
        'If InStr(1, Description(UBound(Description)), "<SPAN") <> 0 Then
        Description(UBound(Description)) = GetLeftWord(ResultBlock(X), "<SPAN", True)
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

     AddToFinal URL(UBound(URL)), Title(UBound(Title)), Description(UBound(Description)), "Google"
    Next
    'Call WriteHTML
End Sub


Public Sub ParseAltavista(Result As String)
    Dim ResultBlock() As String
    Dim pos As String
    Dim X As Integer

   Dim Title() As String
    Dim URL() As String
    Dim Description() As String
    
    ReDim Preserve Title(0) As String
    ReDim Preserve URL(0) As String
    ReDim Preserve Description(0) As String
    
    ReDim Preserve ResultBlock(0) As String
    pos = 1
    While pos <> 0
        pos = InStr(1, Result, "class=result", vbTextCompare)
        Result = Right(Result, Len(Result) - pos)
        ReDim Preserve ResultBlock(UBound(ResultBlock) + 1) As String
        ResultBlock(UBound(ResultBlock)) = GetLeftWord(Result, "</TD>", False)
        Result = Right(Result, Len(Result) - Len(ResultBlock(UBound(ResultBlock))))
    Wend
    For X = 1 To UBound(ResultBlock) - 1
        ReDim Preserve URL(UBound(URL) + 1) As String
        pos = InStr(1, ResultBlock(X), "status='", vbTextCompare)
        If pos <> 0 Then pos = pos + 7
        ResultBlock(X) = Right(ResultBlock(X), Len(ResultBlock(X)) - pos)
        pos = InStr(1, ResultBlock(X), "'", vbTextCompare)
        If pos <> 0 Then URL(UBound(URL)) = Left$(ResultBlock(X), pos - 1)
        
        ResultBlock(X) = Right(ResultBlock(X), Len(ResultBlock(X)) - pos)

        ReDim Preserve Title(UBound(Title) + 1) As String
        Title(UBound(Title)) = GetLeftWord(ResultBlock(X), "</A>", False)
        ResultBlock(X) = Right(ResultBlock(X), Len(ResultBlock(X)) - Len(Title(UBound(Title))))
        Title(UBound(Title)) = GetRightWord(Title(UBound(Title)), Chr(34) & ">", False)

        Title(UBound(Title)) = Replace(Title(UBound(Title)), vbCrLf, "", , , vbTextCompare)

        Title(UBound(Title)) = Replace(Title(UBound(Title)), "  ", " ")
        Title(UBound(Title)) = Replace(Title(UBound(Title)), "  ", " ")
        Title(UBound(Title)) = Replace(Title(UBound(Title)), "  ", " ")
        Title(UBound(Title)) = Replace(Title(UBound(Title)), "  ", " ")

        Title(UBound(Title)) = Replace(Title(UBound(Title)), ">", "", , , vbTextCompare)
        Title(UBound(Title)) = Replace(Title(UBound(Title)), "&amp;", "", , , vbTextCompare)
        If Left$(Title(UBound(Title)), 24) = "http://www.altavista.com" Then Title(UBound(Title)) = "No Title"
Title(UBound(Title)) = Title(UBound(Title)) '& " - From Altavista"

        ReDim Preserve Description(UBound(Description) + 1) As String
        ResultBlock(X) = Replace(ResultBlock(X), vbCrLf, "", , , vbTextCompare)
        ResultBlock(X) = Replace(ResultBlock(X), "  ", " ")
        ResultBlock(X) = Replace(ResultBlock(X), "  ", " ")
        ResultBlock(X) = Replace(ResultBlock(X), "  ", " ")


        pos = InStr(1, ResultBlock(X), "<SPAN class=s>", vbTextCompare)
        If pos <> 0 Then pos = pos + 13

        ResultBlock(X) = Right(ResultBlock(X), Len(ResultBlock(X)) - pos)
        Description(UBound(Description)) = GetLeftWord(ResultBlock(X), "<SPAN", False)
        Description(UBound(Description)) = Replace(Description(UBound(Description)), "<B>", "", , , vbTextCompare)
        Description(UBound(Description)) = Replace(Description(UBound(Description)), "</B>", "", , , vbTextCompare)
        Description(UBound(Description)) = Replace(Description(UBound(Description)), "<BR>", "", , , vbTextCompare)
        Description(UBound(Description)) = Replace(Description(UBound(Description)), "</A>", "", , , vbTextCompare)
        Description(UBound(Description)) = Replace(Description(UBound(Description)), ">", "", , , vbTextCompare)
        Description(UBound(Description)) = Replace(Description(UBound(Description)), "&amp;", "", , , vbTextCompare)


        If Trim(Description(UBound(Description))) = "" Then Description(UBound(Description)) = "No Description"
AddToFinal URL(UBound(URL)), Title(UBound(Title)), Description(UBound(Description)), "Altavista"
        
    Next
    'Call WriteHTML
End Sub

Public Sub ParseJeeves(Result As String)
    On Error Resume Next
    Dim ResultBlock() As String
    Dim pos As String
    Dim X As Integer
   Dim Title() As String
    Dim URL() As String
    Dim Description() As String
    
    ReDim Preserve Title(0) As String
    ReDim Preserve URL(0) As String
    ReDim Preserve Description(0) As String
    
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
    For X = 1 To UBound(ResultBlock) - 1
        ReDim Preserve URL(UBound(URL) + 1) As String
        'Pos = InStr(1, ResultBlock(X), Chr(34))
        URL(UBound(URL)) = GetLeftWord(ResultBlock(X), "return ss('go to", False)
        ResultBlock(X) = Right(ResultBlock(X), Len(ResultBlock(X)) - Len(URL(UBound(URL))))
        URL(UBound(URL)) = GetLeftWord(ResultBlock(X), "')", False)
        ResultBlock(X) = Right(ResultBlock(X), Len(ResultBlock(X)) - Len(URL(UBound(URL))))
        URL(UBound(URL)) = GetRightWord(URL(UBound(URL)), "to ", False)
        URL(UBound(URL)) = Right(URL(UBound(URL)), Len(URL(UBound(URL))) - 2)
        ResultBlock(X) = Right(ResultBlock(X), Len(ResultBlock(X)) - pos)


        ReDim Preserve Title(UBound(Title) + 1) As String
        Title(UBound(Title)) = GetLeftWord(ResultBlock(X), "</A>", False)
        ResultBlock(X) = Right(ResultBlock(X), Len(ResultBlock(X)) - Len(Title(UBound(Title))))
        Title(UBound(Title)) = GetRightWord(Title(UBound(Title)), ">", False)

        Title(UBound(Title)) = Replace(Title(UBound(Title)), vbCrLf, "", , , vbTextCompare)
        Title(UBound(Title)) = Replace(Title(UBound(Title)), "  ", " ")
        Title(UBound(Title)) = Replace(Title(UBound(Title)), "  ", " ")
        Title(UBound(Title)) = Replace(Title(UBound(Title)), "  ", " ")
        Title(UBound(Title)) = Replace(Title(UBound(Title)), "  ", " ")
        Title(UBound(Title)) = Replace(Title(UBound(Title)), "&amp;", "", , , vbTextCompare)
Title(UBound(Title)) = Title(UBound(Title)) '& " - From AskJeeves"


        ReDim Preserve Description(UBound(Description) + 1) As String
        pos = InStr(1, ResultBlock(X), "<DIV", vbTextCompare)
        If pos <> 0 Then pos = pos + 4

        ResultBlock(X) = Right(ResultBlock(X), Len(ResultBlock(X)) - pos)

        Description(UBound(Description)) = GetLeftWord(ResultBlock(X), "</DIV>", False)



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

        AddToFinal URL(UBound(URL)), Title(UBound(Title)), Description(UBound(Description)), "AskJeeves"
    Next
    'Call WriteHTML
End Sub
Public Sub ParseAlltheWeb(Result As String)
    Dim ResultBlock() As String
    Dim pos As String
    Dim X As Integer
       Dim Title() As String
    Dim URL() As String
    Dim Description() As String
    
    ReDim Preserve Title(0) As String
    ReDim Preserve URL(0) As String
    ReDim Preserve Description(0) As String
    
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
    For X = 1 To UBound(ResultBlock) - 1
        ReDim Preserve URL(UBound(URL) + 1) As String
        pos = InStr(1, ResultBlock(X), "class=resURL", vbTextCompare)
        If pos <> 0 Then pos = pos + 11
        
        URL(UBound(URL)) = Right(ResultBlock(X), Len(ResultBlock(X)) - pos)
        'pos = InStr(1, ResultBlock(x), "'", vbTextCompare)
        'URL(UBound(URL)) = Left$(ResultBlock(x), pos - 1)
        URL(UBound(URL)) = GetLeftWord(URL(UBound(URL)), "</span>", False)
       URL(UBound(URL)) = Replace(URL(UBound(URL)), ">", "")
        'ResultBlock(x) = Right(ResultBlock(x), Len(ResultBlock(x)) - pos)

        ReDim Preserve Title(UBound(Title) + 1) As String
        Title(UBound(Title)) = GetLeftWord(ResultBlock(X), "</A>", False)
        Title(UBound(Title)) = GetRightWord(Title(UBound(Title)), "href=", False)
        ResultBlock(X) = Right(ResultBlock(X), Len(ResultBlock(X)) - Len(Title(UBound(Title))))
        Title(UBound(Title)) = Right(Title(UBound(Title)), Len(Title(UBound(Title))) - Len(GetLeftWord(Title(UBound(Title)), ">", False)))
        'Title(UBound(Title)) = GetRightWord(Title(UBound(Title)), ">", True)
        Title(UBound(Title)) = Replace(Title(UBound(Title)), vbCrLf, "", , , vbTextCompare)
        Title(UBound(Title)) = Replace(Title(UBound(Title)), "<SPAN class=hlight>", "", , , vbTextCompare)
        Title(UBound(Title)) = Replace(Title(UBound(Title)), "</SPAN>", "", , , vbTextCompare)
        Title(UBound(Title)) = Replace(Title(UBound(Title)), ">", "", , , vbTextCompare)
    
Title(UBound(Title)) = Title(UBound(Title)) '& " - From AlltheWeb"

                ReDim Preserve Description(UBound(Description) + 1) As String
pos = InStr(1, ResultBlock(X), "<br", vbTextCompare)
If pos <> 0 Then ResultBlock(X) = Right(ResultBlock(X), Len(ResultBlock(X)) - pos - 2)
        Description(UBound(Description)) = GetLeftWord(ResultBlock(X), "<br />", False)
        

'MsgBox Description(UBound(Description))
        Description(UBound(Description)) = Replace(Description(UBound(Description)), vbCrLf, "", , , vbTextCompare)
        Description(UBound(Description)) = Replace(Description(UBound(Description)), "/>", "", , , vbTextCompare)
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

    
        
        AddToFinal URL(UBound(URL)), Title(UBound(Title)), Description(UBound(Description)), "AllTheWeb"

    Next
    'Call WriteHTML
End Sub

Public Sub ParseTeoma(Result As String)
    Dim ResultBlock() As String
    Dim pos As String
    Dim X As Integer
   Dim Title() As String
    Dim URL() As String
    Dim Description() As String
    
    ReDim Preserve Title(0) As String
    ReDim Preserve URL(0) As String
    ReDim Preserve Description(0) As String
    
    ReDim Preserve ResultBlock(0) As String
    pos = 1
    While pos <> 0
        pos = InStr(1, Result, "<DIV><SPAN class=resultTxt>", vbTextCompare)
        Result = Right(Result, Len(Result) - pos)
        ReDim Preserve ResultBlock(UBound(ResultBlock) + 1) As String
        ResultBlock(UBound(ResultBlock)) = GetLeftWord(Result, "</DIV>", False)
        Result = Right(Result, Len(Result) - Len(ResultBlock(UBound(ResultBlock))))
    Wend
    For X = 1 To UBound(ResultBlock) - 1
        ReDim Preserve URL(UBound(URL) + 1) As String
        pos = InStr(1, ResultBlock(X), Chr(34) & ">", vbTextCompare)
        URL(UBound(URL)) = Left$(ResultBlock(X), pos - 1)
        ResultBlock(X) = Right(ResultBlock(X), Len(ResultBlock(X)) - pos)

        URL(UBound(URL)) = GetRightWord(URL(UBound(URL)), "u=", False)
        URL(UBound(URL)) = Right(URL(UBound(URL)), Len(URL(UBound(URL))) - 1)

        ReDim Preserve Title(UBound(Title) + 1) As String
        Title(UBound(Title)) = GetLeftWord(ResultBlock(X), "</A>", False)
        ResultBlock(X) = Right(ResultBlock(X), Len(ResultBlock(X)) - Len(Title(UBound(Title))))
        Title(UBound(Title)) = GetRightWord(Title(UBound(Title)), Chr(34), True)
        Title(UBound(Title)) = Replace(Title(UBound(Title)), vbCrLf, "", , , vbTextCompare)
        Title(UBound(Title)) = Replace(Title(UBound(Title)), "<SPAN class=bold", "", , , vbTextCompare)

        'Title(UBound(Title)) = Replace(Title(UBound(Title)), "<SPAN","", , , vbTextCompare)
        Title(UBound(Title)) = Replace(Title(UBound(Title)), "</SPAN>", "", , , vbTextCompare)
        Title(UBound(Title)) = Replace(Title(UBound(Title)), ">", "", , , vbTextCompare)

Title(UBound(Title)) = Title(UBound(Title)) '& " - From Teoma"

        ReDim Preserve Description(UBound(Description) + 1) As String

        Description(UBound(Description)) = GetLeftWord(ResultBlock(X), "<A", False)
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
AddToFinal URL(UBound(URL)), Title(UBound(Title)), Description(UBound(Description)), "Teoma"
       
    Next
    'Call WriteHTML
End Sub
Public Sub ParseDMOZ(Result As String)
    Dim ResultBlock() As String
    Dim pos As String
    Dim X As Integer
   Dim Title() As String
    Dim URL() As String
    Dim Description() As String
    
    ReDim Preserve Title(0) As String
    ReDim Preserve URL(0) As String
    ReDim Preserve Description(0) As String
    
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
    For X = 1 To UBound(ResultBlock) - 1
        ReDim Preserve URL(UBound(URL) + 1) As String
        pos = InStr(1, ResultBlock(X), Chr(34), vbTextCompare)
        ResultBlock(X) = Right(ResultBlock(X), Len(ResultBlock(X)) - pos)
        pos = InStr(1, ResultBlock(X), Chr(34), vbTextCompare)
        If pos <> 0 Then
        URL(UBound(URL)) = Left$(ResultBlock(X), pos - 1)
        End If
        ResultBlock(X) = Right(ResultBlock(X), Len(ResultBlock(X)) - pos)


        ReDim Preserve Title(UBound(Title) + 1) As String
        Title(UBound(Title)) = GetLeftWord(ResultBlock(X), "</A>", False)
        ResultBlock(X) = Right(ResultBlock(X), Len(ResultBlock(X)) - Len(Title(UBound(Title))))

        Title(UBound(Title)) = Replace(Title(UBound(Title)), vbCrLf, "", , , vbTextCompare)


        Title(UBound(Title)) = Replace(Title(UBound(Title)), "<B>", "", , , vbTextCompare)
        Title(UBound(Title)) = Replace(Title(UBound(Title)), "</B>", "", , , vbTextCompare)
        Title(UBound(Title)) = Replace(Title(UBound(Title)), ">", "", , , vbTextCompare)
        Title(UBound(Title)) = Replace(Title(UBound(Title)), "  ", " ")
        Title(UBound(Title)) = Replace(Title(UBound(Title)), "  ", " ")
        Title(UBound(Title)) = Replace(Title(UBound(Title)), "  ", " ")
Title(UBound(Title)) = Title(UBound(Title)) '& " - From Open Directory"


        ReDim Preserve Description(UBound(Description) + 1) As String

        Description(UBound(Description)) = GetLeftWord(ResultBlock(X), "<SMALL>", False)
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
AddToFinal URL(UBound(URL)), Title(UBound(Title)), Description(UBound(Description)), "DMOZ"
      
    Next
    'Call WriteHTML
End Sub
