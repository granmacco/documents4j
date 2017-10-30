' Configure error handling to jump to next line.
On Error Resume Next

' We don't want the Excel Conversor up, so let's always answer that there is no Excel instance
WScript.Quit -6

' Try to get running MS Excel instance.
' Dim excelApplication
' Set excelApplication = GetObject(, "Excel.Application")

' Signal whether or not such an instance could not be found.
' If Err <> 0 then
  ' WScript.Quit -6
' Else
  ' WScript.Quit 3
' End If
