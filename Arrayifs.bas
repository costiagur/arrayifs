Attribute VB_Name = "Arrayifs"
Function Arrayifs(valrange, testrange1, condition1, Optional testrange2 = "", Optional condition2 = "", Optional testrange3 = "", Optional condition3 = "", Optional testrange4 = "", Optional condition4 = "", Optional delimiter As String = "", Optional uniques As Boolean = False)

Dim i As Integer, j As Integer
Dim testrangearr(), conditionarr(), checkstr As String, condarr(4)
Dim reslist(), midval, midcond, removed As String
Dim valarr, resitem, newarr
Dim dictuniq As Object

ReDim reslist(valrange.Count)
ReDim conditionarr(4)
ReDim testrangearr(4)

valarr = valrange

condarr(0) = condition1
condarr(1) = condition2
condarr(2) = condition3
condarr(3) = condition4

i = 0
j = 0
For i = 0 To 3
    If condarr(i) <> "" Then
        
        midcond = Replace(condarr(i), "<", "")
        midcond = Replace(midcond, ">", "")
        midcond = Replace(midcond, "=", "")
            
        removed = Mid(condarr(i), 1, Len(condarr(i)) - Len(midcond)) 'get operator
            
        If IsDate(midcond) Then midcond = Val(Format(midcond, "General Number")) 'convert date to number
            
        If Not IsNumeric(midcond) Then midcond = Chr(34) & midcond & Chr(34)
        
        If removed = "" Then
            midcond = "=" & midcond
        Else
            midcond = removed & midcond
        End If
        
        conditionarr(j) = midcond
        
        Select Case i
            Case 0
                testrangearr(j) = testrange1.Value
            Case 1
                testrangearr(j) = testrange2.Value
            Case 2
                testrangearr(j) = testrange3.Value
            Case 3
                testrangearr(j) = testrange4.Value
        End Select
        
        j = j + 1
    End If

Next i

ReDim Preserve conditionarr(j - 1)
ReDim Preserve testrangearr(j - 1)

i = 1

For i = LBound(valarr) To UBound(valarr)
    
    If IsEmpty(valarr(i, 1)) Then GoTo nexti
    
    checkstr = ""

    j = 0

    For j = 0 To UBound(conditionarr)
        
        midval = testrangearr(j)(i, 1)
        
        If IsEmpty(midval) Then GoTo nextj
        
        If IsDate(midval) Then midval = Val(Format(midval, "General Number"))
        
        If Not IsNumeric(midval) Then midval = Chr(34) & midval & Chr(34)
        
        checkstr = checkstr & midval & conditionarr(j) & "," 'create evaluation string

nextj:
    Next j
    
    checkstr = "and(" & checkstr & "TRUE)" 'if missing all the conditions the result is true
   
    If Evaluate(checkstr) = True Then reslist(i - 1) = valarr(i, 1)
    
nexti:
Next i

ReDim Preserve reslist(i - 2)

If uniques = True Then
    Set dictuniq = CreateObject("Scripting.Dictionary")
    i = 0
    
    For i = 0 To UBound(reslist)
        On Error Resume Next
        dictuniq.Add Key:=reslist(i), Item:=1
    Next i
        
    If dictuniq.exists("") Then dictuniq.Remove ""
    
    If delimiter <> "" Then
        Arrayifs = Trim(Join(dictuniq.keys, delimiter))
    Else
        Arrayifs = dictuniq.keys
    End If
    
    Set dictuniq = Nothing
    
ElseIf uniques = False Then
    If delimiter <> "" Then
        Arrayifs = Trim(Join(reslist, delimiter))
    Else
        Arrayifs = reslist
    End If
End If
    
End Function
