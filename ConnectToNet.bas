Attribute VB_Name = "ConnectToNet"
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Function Run(strFilePath As String, Optional strParms As String, Optional strDir As String) As String
       
  Const SW_SHOW = 5
  
  'Run the Program and Evaluate errors
  Select Case ShellExecute(0, "Open", strFilePath, strParms, strDir, SW_SHOW)
  Case 0
    Run = "Insufficent system memory or corrupt program file"
  Case Else
    Run = ""
  End Select

End Function
