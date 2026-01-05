Attribute VB_Name = "Library"
Public Enum AREAS
    MARKET_ONE_BOX_ONE
End Enum

Public Function Add(ByRef ListArray() As Variant, insert As Variant)
    If IsEmpty(ListArray(0)) Then
        ListArray(0) = insert
    Else
        ReDim Preserve ListArray(UBound(ListArray) + 1)
        ListArray(UBound(ListArray)) = insert
    End If
End Function

Public Function AddNewNameToEmployees(ByRef ListArray() As Variant, insert As String)
    Set Lab = Worksheets("Presentation-Lab")

    Lowestpoint = "A" + CStr(Lab.Cells(Rows.Count, "A").End(xlUp).Row + 1)
    Lab.Range(Lowestpoint) = insert
    ' CStr(mycell.Value)
    Call Add(ListArray, insert)
End Function


Function returnAreaInLab(target As Integer) As String

' TODO Change this to ENUMS as the INTS don't tell enough + change to Switch handler
' 0 --Labtab Employee Area, box 3
' 1 --Labtab Last spot for employees, box 3

If target = 0 Then
    If (Worksheets("Presentation-Lab").Cells(Rows.Count, "A").End(xlUp).Row < 27) Then
        returnAreaInLab = "A27:A27"
    Else
        returnAreaInLab = "A27:A" + CStr(Worksheets("Presentation-Lab").Cells(Rows.Count, "A").End(xlUp).Row)
    End If
ElseIf target = 1 Then
    returnAreaInLab = "Null"
End If


End Function


Function returnAreaInExcell(target As Integer, Worklist As Worksheet) As String

' TODO Change this to ENUMS as the INTS don't tell enough + change to Switch handler
' --Worklist Employee Area, box 3

If target = 0 Then
    If (Worklist.Cells(Rows.Count, "H").End(xlUp).Row < 1) Then
        returnAreaInExcell = "H2:H2"
    Else
        returnAreaInExcell = "H2:H" + CStr(Worklist.Cells(Rows.Count, "H").End(xlUp).Row)
    End If
ElseIf target = 1 Then
    returnAreaInExcell = "Null"
End If


End Function

Sub Test_Hi_Function()

    Dim TstHiFunc As String
    
    ' send "Hello World" to Function Hi as a parameter
    ' TstHiFunc gets the returned string result
    TstHiFunc = returnAreaInExcell(0, Worksheets("NL Worklist"))
    
    ' for debug only
    MsgBox TstHiFunc

End Sub

