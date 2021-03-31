Attribute VB_Name = "Module1"
Option Explicit

Function Arrayifs(valrange, testrange1, condition1, Optional testrange2 = "", Optional condition2 = "", Optional testrange3 = "", Optional condition3 = "", Optional testrange4 = "", Optional condition4 = "", Optional asstring As Boolean = False)

Dim i As Integer, j As Integer
Dim testrangearr As New Collection, conditionarr As New Collection, checkstr As String, condarr As New Collection, testarr As New Collection
Dim reslist As New Collection, resarr(), midval, midcond, removed As String

condarr.Add condition1
condarr.Add condition2
condarr.Add condition3
condarr.Add condition4

testarr.Add testrange1
testarr.Add testrange2
testarr.Add testrange3
testarr.Add testrange4

For i = 1 To condarr.Count

    If condarr.Item(i) <> "" Then
        
        midcond = Replace(condarr.Item(i), "<", "")
        midcond = Replace(midcond, ">", "")
        midcond = Replace(midcond, "=", "")
            
        removed = mid(condarr.Item(i), 1, Len(condarr.Item(i)) - Len(midcond)) 'get operator
            
        If IsDate(midcond) Then midcond = Val(Format(midcond, "General Number")) 'convert date to number
            
        If Not IsNumeric(midcond) Then midcond = Chr(34) & midcond & Chr(34)
        
        If removed = "" Then
            midcond = "=" & midcond
        Else
            midcond = removed & midcond
        End If
        
        conditionarr.Add midcond
                
        'Debug.Print conditionarr(conditionarr.Count)
        
        testrangearr.Add testarr.Item(i)
    End If

Next i

Set condarr = Nothing
Set testarr = Nothing

For i = 1 To valrange.Count
    
    If IsEmpty(valrange(i).Value2) Then GoTo nexti
    
    checkstr = ""

    For j = 1 To conditionarr.Count
        
        midval = testrangearr.Item(j)(i).Value2
        
        If IsEmpty(midval) Then GoTo nextj
        
        If IsDate(midval) Then midval = Val(Format(midval, "General Number"))
        
        If Not IsNumeric(midval) Then midval = Chr(34) & midval & Chr(34)
        
        checkstr = checkstr & midval & conditionarr.Item(j) & "," 'create evaluation string

nextj:
    Next j
    
    checkstr = "and(" & checkstr & "TRUE)" 'if missing all the conditions the result is true
    
    'Debug.Print checkstr

    If Evaluate(checkstr) = True Then reslist.Add valrange(i).Value2
    
nexti:
Next i

Set testrangearr = Nothing
Set conditionarr = Nothing

ReDim resarr(reslist.Count)

i = 1

For i = 1 To reslist.Count
    resarr(i - 1) = reslist(i)
Next i

Set reslist = Nothing

'Debug.Print Join(resarr, "_")

If asstring = True Then
    Arrayifs = Join(resarr, "_")
Else
    Arrayifs = resarr
End If


End Function

