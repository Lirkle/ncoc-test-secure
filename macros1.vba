
Option Explicit

Public Sub Color_AllNames_InitialDotSurname()
    Dim targetColor As Long
    targetColor = GetTargetColorRGB()
    If targetColor = -1 Then Exit Sub

    Dim colorLines As Boolean
    colorLines = (MsgBox("Color lines too?", vbYesNo + vbQuestion, "Options") = vbYes)

    Dim sld As Slide
    Dim totalColored As Long, totalFixed As Long

    For Each sld In ActivePresentation.Slides
        totalColored = totalColored + ProcessShapesCollection(sld.Shapes, targetColor, totalFixed, colorLines)
    Next sld

    MsgBox "Done." & vbCrLf & _
           "Fixed spaces after dot: " & totalFixed & vbCrLf & _
           "Colored matches: " & totalColored, vbInformation, "Status"
End Sub

Private Function GetTargetColorRGB() As Long
    On Error GoTo fallback

    ' 1) If text is selected - get its color (you select it on palette beforehand)
    If ActiveWindow.Selection.Type = ppSelectionText Then
        GetTargetColorRGB = ActiveWindow.Selection.TextRange.Font.Color.RGB
        Exit Function
    End If

fallback:
    ' 2) Otherwise ask for HEX
    GetTargetColorRGB = AskColorHexRGB()
End Function

Private Function AskColorHexRGB() As Long
    Dim s As String
    s = InputBox("Enter color in HEX without # (example: 1E90FF). Or leave empty to cancel.", "Color")
    s = Trim$(s)

    If Len(s) = 0 Then
        AskColorHexRGB = -1
        Exit Function
    End If

    s = Replace$(s, "#", "")
    If Len(s) <> 6 Then
        MsgBox "HEX must be exactly 6 characters, example: 1E90FF", vbExclamation
        AskColorHexRGB = -1
        Exit Function
    End If

    On Error GoTo bad
    Dim r As Long, g As Long, b As Long
    r = CLng("&H" & Mid$(s, 1, 2))
    g = CLng("&H" & Mid$(s, 3, 2))
    b = CLng("&H" & Mid$(s, 5, 2))

    AskColorHexRGB = RGB(r, g, b)
    Exit Function

bad:
    MsgBox "Could not read HEX. Correct example: 1E90FF", vbExclamation
    AskColorHexRGB = -1
End Function

Private Function ProcessShapesCollection(ByVal shps As Shapes, ByVal targetColor As Long, ByRef fixedSpaces As Long, ByVal colorLines As Boolean) As Long
    Dim shp As Shape
    Dim cnt As Long

    For Each shp In shps
        cnt = cnt + ProcessShapeRecursive(shp, targetColor, fixedSpaces, colorLines)
    Next shp

    ProcessShapesCollection = cnt
End Function

Private Function ProcessShapeRecursive(ByVal shp As Shape, ByVal targetColor As Long, ByRef fixedSpaces As Long, ByVal colorLines As Boolean) As Long
    Dim cnt As Long
    Dim i As Long

    ' Groups
    If shp.Type = msoGroup Then
        For i = 1 To shp.GroupItems.Count
            cnt = cnt + ProcessShapeRecursive(shp.GroupItems(i), targetColor, fixedSpaces, colorLines)
        Next i
        ProcessShapeRecursive = cnt
        Exit Function
    End If

    If colorLines Then
        Dim isConn As Boolean
        isConn = False

        On Error Resume Next
        ' If ConnectorFormat is available - this is a connector (even if not connected at ends)
        Dim cType As Long
        cType = shp.ConnectorFormat.Type
        isConn = (Err.Number = 0)
        Err.Clear
        On Error GoTo 0

        If isConn Then
            If shp.Line.Visible = msoTrue Then
                shp.Line.ForeColor.RGB = targetColor
            End If
        End If
    End If

    On Error Resume Next

    ' Tables
    If shp.HasTable Then
        Dim r As Long, c As Long
        For r = 1 To shp.Table.Rows.Count
            For c = 1 To shp.Table.Columns.Count
                cnt = cnt + ProcessTextRange(shp.Table.Cell(r, c).Shape.TextFrame.TextRange, targetColor, fixedSpaces, colorLines)
            Next c
        Next r
        ProcessShapeRecursive = cnt
        Exit Function
    End If

    ' Regular text
    If shp.HasTextFrame Then
        If shp.TextFrame.HasText Then
            cnt = cnt + ProcessTextRange(shp.TextFrame.TextRange, targetColor, fixedSpaces, colorLines)
        End If
    End If

    ProcessShapeRecursive = cnt
End Function

Private Function ProcessTextRange(ByVal tr As TextRange, ByVal targetColor As Long, ByRef fixedSpaces As Long, ByVal colorLines As Boolean) As Long
    On Error GoTo EH

    Dim text As String
    text = tr.text
    If Len(text) = 0 Then Exit Function

    ' 1) Remove spaces after dot only in constructions:
    '    1-2 letters + "." + spaces + Capital letter (start of surname)
    Dim fixed As Long
    fixed = NormalizeSpacesAfterDot(tr)
    fixedSpaces = fixedSpaces + fixed

    ' 2) Color all matches: 1-2 letters + "." + Surname (with capital letter)
    ProcessTextRange = ColorAllMatches(tr, targetColor)
    Exit Function

EH:
    ProcessTextRange = 0
End Function

Private Function NormalizeSpacesAfterDot(ByVal tr As TextRange) As Long
    Dim original As String, updated As String
    original = tr.text

    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Global = True
    re.MultiLine = True
    re.IgnoreCase = False

    ' Example: "I.   Partilova" -> "I.Partilova"
    re.Pattern = "([A-Za-zА-Яа-яЁё]{1,2})\.\s+([A-ZА-ЯЁ])"

    updated = re.Replace(original, "$1.$2")

    If updated <> original Then
        ' Count the number of fixes
        Dim m As Object, ms As Object
        Set ms = re.Execute(original)
        NormalizeSpacesAfterDot = ms.Count

        tr.text = updated
    Else
        NormalizeSpacesAfterDot = 0
    End If
End Function

Private Function ColorAllMatches(ByVal tr As TextRange, ByVal targetColor As Long) As Long
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Global = True
    re.MultiLine = True
    re.IgnoreCase = False

    ' Matches like: "I.Partilova", "AB.Ivanov", also inside "I.Partilova/A.Ivanov"
    ' Require capital letter at start of surname, then letters/hyphen/apostrophe
    re.Pattern = "([A-Za-zА-Яа-яЁё]{1,2})\.([A-ZА-ЯЁ][A-Za-zА-Яа-яЁё'\-]*)"

    Dim s As String
    s = tr.text

    Dim matches As Object, k As Long
    Set matches = re.Execute(s)

    ' Important: color from end to avoid breaking indices
    For k = matches.Count - 1 To 0 Step -1
        Dim startPos As Long, ln As Long
        startPos = matches(k).FirstIndex + 1        ' TextRange.Characters 1-based
        ln = matches(k).Length

        With tr.Characters(startPos, ln).Font
            .Color.RGB = targetColor
        End With
    Next k
    
    ' INSIDE ColorAllMatches, after Next k

Dim reSlash As Object, msSlash As Object, j As Long
Set reSlash = CreateObject("VBScript.RegExp")
reSlash.Global = True
reSlash.MultiLine = True
reSlash.IgnoreCase = False

reSlash.Pattern = "([A-Za-zА-Яа-яЁё]{1,2}\.[A-ZА-ЯЁ][A-Za-zА-Яа-яЁё'\-]*)(\s*/\s*)([A-Za-zА-Яа-яЁё]{1,2}\.[A-ZА-ЯЁ][A-Za-zА-Яа-яЁё'\-]*)"


Set msSlash = reSlash.Execute(s)

For j = msSlash.Count - 1 To 0 Step -1
    Dim slashPos As Long
    ' FirstIndex is 0-based, Characters is 1-based
slashPos = msSlash(j).FirstIndex + Len(msSlash(j).SubMatches(0)) + InStr(1, msSlash(j).SubMatches(1), "/", vbBinaryCompare)

    tr.Characters(slashPos, 1).Font.Color.RGB = targetColor
Next j

' At the END of ColorAllMatches don't forget:
ColorAllMatches = matches.Count


    ColorAllMatches = matches.Count
End Function


