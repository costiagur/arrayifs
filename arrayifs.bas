Attribute VB_Name = "Module1"
Option Explicit

Function Arrayifs(valrange, testrange1, condition1, Optional testrange2 = "", Optional condition2 = "", Optional testrange3 = "", Optional condition3 = "", Optional testrange4 = "", Optional condition4 = "", Optional delimiter As String = "")

Dim i As Integer, j As Integer
Dim testrangearr As New Collection, conditionarr As New Collection, checkstr As String, condarr As New Collection
Dim reslist As Object, midval, midcond, removed As String
Dim cond, valarr

Set reslist = CreateObject("Scripting.Dictionary")

valarr = valrange

condarr.Add condition1
condarr.Add condition2
condarr.Add condition3
condarr.Add condition4

i = 0

For Each cond In condarr
    i = i + 1
    If cond <> "" Then
        
        midcond = Replace(cond, "<", "")
        midcond = Replace(midcond, ">", "")
        midcond = Replace(midcond, "=", "")
            
        removed = Mid(cond, 1, Len(cond) - Len(midcond)) 'get operator
            
        If IsDate(midcond) Then midcond = val(format(midcond, "General Number")) 'convert date to number
            
        If Not IsNumeric(midcond) Then midcond = Chr(34) & midcond & Chr(34)
        
        If removed = "" Then
            midcond = "=" & midcond
        Else
            midcond = removed & midcond
        End If
        
        conditionarr.Add midcond
        
        Select Case i
            Case 1
                testrangearr.Add testrange1.Value
            Case 2
                testrangearr.Add testrange2.Value
            Case 3
                testrangearr.Add testrange3.Value
            Case 4
                testrangearr.Add testrange4.Value
        End Select
    End If

Next

Set condarr = Nothing

i = 1

For i = LBound(valarr) To UBound(valarr)
    
    If IsEmpty(valarr(i, 1)) Then GoTo nexti
    
    checkstr = ""

    j = 0

    For Each cond In conditionarr
        j = j + 1
        
        midval = testrangearr.Item(j)(i, 1)
        
        If IsEmpty(midval) Then GoTo nextj
        
        If IsDate(midval) Then midval = val(format(midval, "General Number"))
        
        If Not IsNumeric(midval) Then midval = Chr(34) & midval & Chr(34)
        
        checkstr = checkstr & midval & cond & "," 'create evaluation string

nextj:
    Next cond
    
    checkstr = "and(" & checkstr & "TRUE)" 'if missing all the conditions the result is true

    If Evaluate(checkstr) = True Then reslist.Add i, valarr(i, 1)
    
nexti:
Next i

Set testrangearr = Nothing
Set conditionarr = Nothing

If delimiter <> "" Then
    Arrayifs = Trim(Join(reslist.items, delimiter))
Else
    Arrayifs = reslist.items
End If

End Function

