Attribute VB_Name = "FillExcel"
Public Sub FillExcel()
'
' Macro2 Macro
'

'

    Dim mycell                      As Range
    Dim myrange                     As Variant
    Dim arrayOfEmployees            As Variant
    Dim arrayOfAllocationOrder()    As Variant
    Dim arrayForTransition()        As Variant
    Dim Lab                         As Worksheet
    Dim WorklistNL                  As Worksheet
    Dim SpotIsOccupied              As Boolean
    Dim Lowestpoint                 As String
    Dim UndefinedEmployeesCount     As Long
    Dim PositionToView              As String
    
    
    
    Set Lab = Worksheets("Presentation-Lab")
    Set WorklistNL = Worksheets("NL Worklist")
    
    ' // Below you must share the area with all the employees
    rangeofemployees = "A27:E45" ' + CStr((Lab.Cells(Rows.Count, "D").End(xlUp).Row) + 2)
    arrayOfEmployees = Lab.Range(rangeofemployees).Value
    ' // Below you must share the area where the systems names should check
    RangeOfWhoWorksCases = "H2:H" + CStr(WorklistNL.Cells(Rows.Count, "H").End(xlUp).Row)
    ' Debug.Print (RangeOfWhoWorksCases)
    Set myrange = Worksheets("NL Worklist").Range(RangeOfWhoWorksCases)
    ReDim Preserve arrayOfAllocationOrder(0)
    Dim i As Long
    Dim x As Long
    Dim y As Long
    
    ' // Creates the allocation order found in the available sets
    For i = 0 To 12                                                             '// Run the check 12x as this is the minimum case requirement + case index = i
        For x = 1 To UBound(arrayOfEmployees, 1)                                '// Run the check for every employee in EmployeeRange
            If Not IsEmpty(arrayOfEmployees(x, 1)) Then                         '// Skip the cell if the employee name is emtpy
                If i = arrayOfEmployees(x, 5) Then                              '// Check the working cases matching the current case index.
                    If arrayOfEmployees(x, 5) < arrayOfEmployees(x, 2) Then
                        Debug.Print ("Employees " + CStr(i) + " = " + CStr(arrayOfEmployees(x, 5)) + " -- " + CStr(arrayOfEmployees(x, 2)))
                        'Debug.Print ("Employees a " + CStr(arrayOfEmployees(x, 2)))
                        Call Add(arrayOfAllocationOrder, x)                     '// Add the employee to the allocation order
                        arrayOfEmployees(x, 5) = arrayOfEmployees(x, 5) + 1     '// Add a point to the working cases so
                    End If
                End If
            End If
        Next x
    Next i

' FillTheExcel with the data
    y = 0
    For i = 1 To WorklistNL.Cells(Rows.Count, "A").End(xlUp).Row
        If y < UBound(arrayOfAllocationOrder, 1) Then
            If IsEmpty(WorklistNL.Cells(i, "H").Value) Then
                WorklistNL.Cells(i, "H").Value = CStr(arrayOfEmployees(arrayOfAllocationOrder(y), 1))
            
                y = y + 1
            End If
        End If
    Next i
End Sub

