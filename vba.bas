'Sub AutoExec()
'ActiveWindow.View.DisplayPageBoundaries = False
' not understand enough why this raise error : with block ... not set
'End Sub
Sub AutoOpen()
'Debug.Print ActiveDocument.ActiveWindow.Caption

ActiveDocument.ActiveWindow.View.DisplayPageBoundaries = False
' must use ActiveDocument, or it doesn;t know and throw error Object or With block variable not set ..
End Sub

Sub WaitForWindow()
    Dim WindowTitle As String
    Dim MaxWait As Double
    Dim StartTime As Double
    
    ' Set the window name you are waiting for
    WindowTitle = "Your Window Title Here"
    
    ' Set the maximum time to wait in seconds (10 seconds here)
    MaxWait = 10
    StartTime = Timer
    
    ' Wait until the window exists or timeout is reached
    Do While FindWindow(vbNullString, WindowTitle) = 0
        DoEvents ' Keeps Word responsive and prevents freezing
        
        ' Exit if it takes too long
        If Timer - StartTime > MaxWait Then
            MsgBox "Timed out waiting for the window."
            Exit Sub
        End If
    Loop
    
    ' Code continues here once the window is found
    MsgBox ActiveDocument.ActiveWindow.Caption
End Sub


Sub add_bookmark_here()
'
' bookmark Macro
'
'
    With ActiveDocument.Bookmarks
        .Add Range:=Selection.Range, Name:="hehe"
        .DefaultSorting = wdSortByName
        .ShowHidden = False
    End With
End Sub

Sub go_end()
'
' go_end Macro
'
'
    add_bookmark_here
    Selection.EndKey Unit:=wdStory
    Selection.TypeParagraph
End Sub
Sub go_home()
'
' go_home Macro
'
'
    add_bookmark_here
    Selection.HomeKey Unit:=wdStory
    Selection.TypeParagraph
    Selection.MoveUp Unit:=wdLine, Count:=1
End Sub


Sub bookmark_up()
'
' bookmark_up Macro
'
'
    Selection.GoTo What:=wdGoToBookmark, Name:="hehe"
    Selection.Find.ClearFormatting
    With Selection.Find
        .Text = ""
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
End Sub
Function Clipboard$(Optional s$)
    Dim v: v = s  'Cast to variant for 64-bit VBA support
    With CreateObject("htmlfile")
    With .parentWindow.clipboardData
        Select Case True
            Case Len(s): .setData "text", v
            Case Else:   Clipboard = .getData("text")
        End Select
    End With
    End With
End Function

Function removeclipclip()
'
' removeclipclip Macro
'
'
    Dim hihi As String
    
    hihi = Format(Now(), "yyyy-MM-dd hh.mm.ss")
'    Debug.Print hihi
    Dim MyValue As Integer
    MyValue = Int((10000 * Rnd) + 1)
    Dim myNumString As String

'    Console.WriteLine (MyValue)
    myNumString = hihi + "  vvvvvvvvvvvvv"
'    myNumStringg = String.Format("{0:00000}", MyValue)
'    Dim myNumStringg As String = $"{MyValue:00000}"

'don't copy text to clipboard like this !
'    Selection.TypeParagraph
'    Selection.TypeText Text:="vvvvvvvvvvvvvvvvvvvvvvvvvvv"
'    Selection.TypeText Text:=MyValue + "vv"
'    Selection.TypeText Text:=myNumString
'    Selection.TypeText Text:=myNumString
'    Selection.HomeKey Unit:=wdLine, Extend:=wdExtend
'    Selection.Copy
'    Selection.Delete Unit:=wdCharacter, count:=1
'    Selection.TypeBackspace
Clipboard (myNumString)
End Function

Sub removeclipclip1()
    Dim data As String
    data = removeclipclip()

End Sub
Sub removeclipclip2()
    Dim data As String
    Dim data2 As String
    data = removeclipclip()
    waitTill = Now() + TimeValue("00:00:02")

While Now() < waitTill
    DoEvents
Wend
    data = removeclipclip()
End Sub

Sub removeclipclip3()
    Dim data As String
    data = removeclipclip()
    waitTill = Now() + TimeValue("00:00:02")

While Now() < waitTill
    DoEvents
Wend
    data = removeclipclip()
        waitTill = Now() + TimeValue("00:00:02")

While Now() < waitTill
    DoEvents
Wend
data = removeclipclip()
End Sub
Sub removeclipclip4()
    Dim data As String
    data = removeclipclip()
    waitTill = Now() + TimeValue("00:00:02")

While Now() < waitTill
    DoEvents
Wend
    data = removeclipclip()
        waitTill = Now() + TimeValue("00:00:02")

While Now() < waitTill
    DoEvents
Wend
data = removeclipclip()
        waitTill = Now() + TimeValue("00:00:02")

While Now() < waitTill
    DoEvents
Wend
data = removeclipclip()
End Sub
Sub insertDatetime()
'
' insertDatetime Macro
' insert current date time as text, no update
'
    Selection.insertDatetime DateTimeFormat:="yyyy-MM-dd", InsertAsField:= _
        False, DateLanguage:=wdEnglishUS, CalendarType:=wdCalendarWestern, _
        InsertAsFullWidth:=False
End Sub

Sub putBookmark()
'
' putBookmark Macro
' create a bookmark and hyperlink it later. To make sure each value number is 2 digits, use 100 or 2100 as the subtractor.
'
Dim rng As Range
 ' Grab current selection
'
Debug.Print Selection.Range.Start
Debug.Print Selection.Range.End
myNumString = "a_" + CStr(2100 - Year(Now)) + CStr(100 - Month(Now)) + CStr(100 - Day(Now)) + "_" + CStr(100 - Hour(Now)) + CStr(100 - Minute(Now)) + CStr(100 - Second(Now))
If Selection.Range.Start = Selection.Range.End Then
    Selection.Expand wdParagraph
    Set rng = Selection.Range
'    Debug.Print "----"
'    Debug.Print Selection.Range.Start
'    Debug.Print Selection.Range.End
'    Debug.Print Selection.Range.Words(4).Text 'assume the reviewed reach only to |9|, if it is |10| then this must by .Words(5)
    Selection.Range.Words(5).Select
    
'    rng.MoveStart Unit:=wdWord, Count:=-6
'    Selection.HomeKey Unit:=wdLine
'    Selection.MoveRight Unit:=wdCharacter, Count:=6, Extend:=wdExtend 'when there is no selected text, the bookmark will be booked at the character 6 from left to right, this is to avoid
    
    With ActiveDocument.Bookmarks
        .Add Range:=Selection.Range, Name:=myNumString
        .DefaultSorting = wdSortByName
        .ShowHidden = False
    End With
Else
    With ActiveDocument.Bookmarks
        .Add Range:=Selection.Range, Name:=myNumString
        .DefaultSorting = wdSortByName
        .ShowHidden = False
    End With
End If
End Sub




Function updateValue(arg1 As String)

Dim txt As String
Dim results As Object
Dim result As Object
Dim r As New RegExp
Dim reviews As Integer
Dim final As String

r.Pattern = "\|(\d{1,3})\|" 'dont care about other group, just want to know in the current text has already existed the number of review or not.
r.Global = True
'txt = "** |1|"
'txt = "** |11| hihi"
txt = arg1
Set results = r.Execute(txt)

If results.Count = 0 Then
Debug.Print "no exiting"
final = "|1|"
Else
reviews = results.item(0).SubMatches.item(0)
reviews = reviews + 1
final = "|" + CStr(reviews) + "|"

End If

updateValue = final
End Function

Function oldValue(arg As String)

Dim txt As String
Dim results As Object
Dim result As Object
Dim r As New RegExp

txt = arg

r.Pattern = "(\s*\**\s*)(\|\d+\|\s*)*"
r.Global = True
Set results = r.Execute(txt) 'get all the matches


'Debug.Print results.item(0).SubMatches.item(1)
'Debug.Print Trim(results.item(0).SubMatches.item(0))


oldValue = results.item(0).SubMatches.item(1)

End Function
Public Function Contains(col As Collection, value As String) As Boolean
'this will use collection to manage objects in vba instead of manage array size.


    For Each item In col
        If value = item Then
        Contains = True
        Exit Function
        Else
        Contains = False
        End If
    Next item

End Function









Sub markReviewed()
    Dim bkm As bookmark
    Dim msg As String
    Dim oldValuedata As String
    Dim updateValuedata As String
    
    ' Expand selection to paragraph
    Selection.Expand wdParagraph
    
        ' Set a range tracking the current selection
    Set selectedRange = Selection.Paragraphs
    paragraphsCount = selectedRange.Count
    
    For Each Paragraph In Selection.Paragraphs
    Debug.Print Paragraph.Range.Text
    
    If Len(Paragraph.Range.Text) > 4 Then
    If updateValue(Paragraph.Range.Text) = "|1|" Then
       Paragraph.Range.InsertBefore "|1| " 'insert before the new mark reviewed
    Else
        Set rngPara = Paragraph.Range
        oldValuedata = oldValue(Paragraph.Range.Text)
        updateValuedata = updateValue(Paragraph.Range.Text)
        
        Debug.Print oldValuedata
        Debug.Print updateValuedata

         With rngPara.Find
             .ClearFormatting ' Clears any previous find formatting
             .Text = oldValuedata ' The text you want to find
             .Replacement.ClearFormatting ' Clears any previous replacement formatting
             .Replacement.Text = updateValuedata & " " ' The text you want to replace with
             .Forward = True ' Search forward
             .Wrap = wdFindStop ' Crucially, this limits the search to the selection
             .Format = False ' Do not consider formatting in the search
             .MatchCase = False ' Do not match case
             .MatchWholeWord = False ' Do not match whole word only
             .Execute Replace:=wdReplaceOne ' Execute the replacement for all occurrences within the selection
         End With
         
    End If
    End If
    Next Paragraph

End Sub

Sub SelectTextBetweenBookmarks()
    Dim rngEnd As Variant
    Set rngStart = ActiveDocument.Bookmarks("hehe").Range
    Debug.Print rngStart
    rngEnd = Selection.Start
    Debug.Print rngEnd
    If rngEnd > rngStart.Start Then
    ActiveDocument.Range(rngStart.Start, rngEnd).Select
    Else
    ActiveDocument.Range(rngEnd, rngStart.Start).Select
    End If
    
End Sub
