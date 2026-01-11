Attribute VB_Name = "Main"
Sub Main()
Application.ScreenUpdating = False

'
' Countusers Macro
' The functionality to count the users. This way we can go over the worklist and find all the users
'

'
   

Call CheckListforUndefinedEmployees.CheckListforUndefinedEmployees


Call GetPersonalCaseCount.GetPersonalCaseCount

Call FillExcel.FillExcel

Call GetPersonalCaseCount.GetPersonalCaseCount



End Sub

