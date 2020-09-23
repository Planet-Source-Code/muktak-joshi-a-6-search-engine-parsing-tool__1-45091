Attribute VB_Name = "modMyFunctions"
Option Explicit
Public Type POINTAPI
    X As Long
    y As Long
End Type
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public PCDetails() As String

Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long

Public Function GetRightWord(ByRef Sentense As String, ByRef StopChar As String, RemoveIt As Boolean)
    'On Error Resume Next

    Dim pos As Integer
    pos = InStrRev(Sentense, StopChar, -1, vbTextCompare)
    If pos = 0 Then
        GetRightWord = Sentense
    Else
        GetRightWord = Right(Sentense, Len(Sentense) - pos)
    End If

    If RemoveIt = True Then Sentense = Left$(Sentense, pos)

End Function



Public Function GetLeftWord(ByRef Sentense As String, ByRef StopChar As String, RemoveIt As Boolean)
    'On Error Resume Next
    Dim pos As Integer
    pos = InStr(1, Sentense, StopChar, vbTextCompare)
    If pos = 0 Then
        GetLeftWord = Sentense
    Else
        GetLeftWord = Left$(Sentense, pos - 1)
    End If
    If Trim(Sentense) = "" Then Exit Function
    If RemoveIt = True Then Sentense = Right(Sentense, Len(Sentense) - pos - Len(StopChar) + 1)

End Function


Public Function fnSort(aSort As Variant, intAsc As Integer) As Variant
Dim intTempStore
Dim i, j
For i = 0 To UBound(aSort) '- 1
For j = i To UBound(aSort)
'Sort Ascending
If intAsc = 1 Then
If aSort(i) > aSort(j) Then
intTempStore = aSort(i)
aSort(i) = aSort(j)
aSort(j) = intTempStore

intTempStore = FinalURL(i)
FinalURL(i) = FinalURL(j)
FinalURL(j) = intTempStore

intTempStore = FinalTitle(i)
FinalTitle(i) = FinalTitle(j)
FinalTitle(j) = intTempStore

intTempStore = FinalDescription(i)
FinalDescription(i) = FinalDescription(j)
FinalDescription(j) = intTempStore

intTempStore = EngineNames(i)
EngineNames(i) = EngineNames(j)
EngineNames(j) = intTempStore

End If 'i > j
'Sort Descending
Else
If aSort(i) < aSort(j) Then

intTempStore = aSort(i)
aSort(i) = aSort(j)
aSort(j) = intTempStore

intTempStore = FinalURL(i)
FinalURL(i) = FinalURL(j)
FinalURL(j) = intTempStore

intTempStore = FinalTitle(i)
FinalTitle(i) = FinalTitle(j)
FinalTitle(j) = intTempStore

intTempStore = FinalDescription(i)
FinalDescription(i) = FinalDescription(j)
FinalDescription(j) = intTempStore

intTempStore = EngineNames(i)
EngineNames(i) = EngineNames(j)
EngineNames(j) = intTempStore

End If 'i < j
End If 'intAsc = 1
Next 'j
Next 'i
fnSort = aSort
End Function

