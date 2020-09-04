Option Explicit

Function Arrayifs(valrange, testrange1, condition1, Optional testrange2 = "", Optional condition2 = "", Optional testrange3 = "", Optional condition3 = "", Optional testrange4 = "", Optional condition4 = "", Optional ifsort As Boolean = True)

Dim i As Integer, j As Integer
Dim checkrangearr, conditionarr, checkstrarr(4) As String, andcond As Boolean
Dim ResList, TheItem, numeric As Boolean

checkrangearr = Array(testrange1, testrange2, testrange3, testrange4)

conditionarr = Array(condition1, condition2, condition3, condition4)
 
Set ResList = CreateObject("System.Collections.ArrayList")

ResList.Clear

numeric = True


For i = 0 To 3

    If VarType(conditionarr(i)) <> 8 Then 'if not string then add "=" sign

        conditionarr(i) = "=" & conditionarr(i)

    End If

Next i


For i = 1 To valrange.count

    andcond = True

    For j = 0 To 3

        If IsObject(checkrangearr(j)) = True Then 'if there is no checkrange, skip this check

            checkstrarr(j) = checkrangearr(j)(i) & conditionarr(j) 'create evaluation string

            andcond = (andcond And Evaluate(checkstrarr(j))) 'combine evaluation

        End If

    Next j
  
    If andcond = True Then
   
        ResList.Add valrange(i).value
       
    End If
    
Next i

If ifsort = True Then
    ResList.Sort
End If

Arrayifs = ResList.ToArray

End Function

