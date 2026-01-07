Attribute VB_Name = "GetPersonalCaseCount"
Public Sub GetPersonalCaseCount()
'
' Check how many cases people worked
' this scripts looks at the point of view of the employee, My name is X and the name i look at is Y, do we match?

'   TODO Make a function that reads the Employees and add them to an Array
'   TODO make a more efficient loop so the system will work faster and more secure

'
Dim mycell                      As Range
Dim myrange                     As Variant
Dim arrayOfEmployees            As Variant
Dim arrayOfUndefinedEmployees() As String
Dim IndividualCells             As Integer
Dim SharedCells                 As Integer
Dim Lab                         As Worksheet
Dim WorklistNL                  As Worksheet
Dim NewName                     As Boolean


Set Lab = Worksheets("Presentation-Lab")
Set WorklistNL = Worksheets("NL Worklist")

' // Below you must share the area with all the employees
rangeofemployees = returnAreaInLab(0)
arrayOfEmployees = Range(rangeofemployees).Value
' // Below you must share the area where the systems names should check
RangeOfEmployeesWorkingCases = "H2:H" + CStr(WorklistNL.Cells(Rows.Count, "H").End(xlUp).Row)
Set myrange = Worksheets("NL Worklist").Range(RangeOfEmployeesWorkingCases)

Dim i As Long
' // This forloop checks in the area how many times a name shown in the employees list has come in
For i = 1 To UBound(arrayOfEmployees, 1) ' LBound(arrayOfEmployees, 1) To UBound(arrayOfEmployees, 1)
    IndividualCells = 0
    SharedCells = 0
    NewName = False
    
    
    ' // This loop checks the cells if the name matches, matches partially or if we need to add the new employee to the list
    For Each mycell In myrange
        If IsEmpty(mycell.Value) Then
            
        Else
            If mycell.Value = arrayOfEmployees(i, 1) Then
                IndividualCells = IndividualCells + 1
            ElseIf InStr(mycell.Value, arrayOfEmployees(i, 1)) > 0 Then
                SharedCells = SharedCells + 1
                Debug.Print ("2 =" + CStr(SharedCells))
            Else

                
            End If
        End If
    Next mycell
    
    If IndividualCells > 0 Then
        returnAreaInLab (0)
        Lab.Cells(26 + i, 3).Value = IndividualCells
    End If
    If SharedCells > 0 Then
        Lab.Cells(26 + i, 4).Value = SharedCells
    End If
Next i



End Sub
