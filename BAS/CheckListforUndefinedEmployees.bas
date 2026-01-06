Attribute VB_Name = "CheckListforUndefinedEmployees"
Public Sub CheckListforUndefinedEmployees()

    Dim mycell                      As Range
    Dim myrange                     As Variant
    Dim arrayOfEmployees            As Variant
    Dim arrayOfUndefinedEmployees() As Variant
    Dim arrayForTransition()        As Variant
    Dim Lab                         As Worksheet
    Dim WorklistNL                  As Worksheet
    Dim EmployeeIsInList            As Boolean
    Dim Lowestpoint                 As String
    Dim UndefinedEmployeesCount     As Long
    
    
    
    Set Lab = Worksheets("Presentation-Lab")
    Set WorklistNL = Worksheets("NL Worklist")
    
    ' // Below you must share the area with all the employees
    rangeofemployees = returnAreaInLab(0)
    arrayOfEmployees = Range(rangeofemployees).Value
    ' // Below you must share the area where the systems names should check
    RangeOfEmployeesWorkingCases = returnAreaInExcell(0, WorklistNL)
    Set myrange = WorklistNL.Range(RangeOfEmployeesWorkingCases)
    ReDim Preserve arrayOfUndefinedEmployees(0)
    Dim i As Long
    Dim x As Long
    Dim y As Long
    
    ' // Go over the worklist names and check if they are already in the list
    For Each mycell In myrange
        EmployeeIsInList = False
        ' Is the cell empty or does the name say Terminated?
         If IsEmpty(mycell.Value) Or mycell.Value = "Terminated" Then
            ' --> Continue no action required
            EmployeeIsInList = True
        Else
            arrayOfEmployees = Range(rangeofemployees).Value
            If (Range(rangeofemployees).Count = 1) Then
                If (IsEmpty(arrayOfEmployees) = True) Then
                    EmployeeIsInList = False
                Else
                    EmployeeIsInList = DoNamesMatch(CStr(mycell.Value), CStr(arrayOfEmployees))
                End If
            Else
                ' Check the list if the cell matches one of the colleagues
                For i = 1 To UBound(arrayOfEmployees, 1)
                    If (DoNamesMatch(CStr(mycell.Value), CStr(arrayOfEmployees(i, 1))) = True) Then
                        EmployeeIsInList = True
                        Exit For
                    End If
                Next i
            End If
        End If
        If EmployeeIsInList = False Then
            Call AddNewNameToEmployees(arrayOfUndefinedEmployees, CStr(mycell.Value))
            rangeofemployees = returnAreaInLab(0)
        End If
    Next mycell
        
End Sub


Function DoNamesMatch(target As String, match As String) As Boolean
        If target = match Then
             ' --> Continue no action required
            DoNamesMatch = True
        ElseIf InStr(target, match) > 0 Then
             ' --> Continue no action required
            DoNamesMatch = True
        Else
           ' --> A new name is found
            DoNamesMatch = False
        End If
End Function
