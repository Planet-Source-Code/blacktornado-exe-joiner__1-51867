Attribute VB_Name = "mdlLoader"
Option Explicit
Dim MyFBF As New clsFileBinder

Sub Main()
On Error Resume Next
MyFBF.OpenPackage IIf(Right(App.Path, "1") = "\", App.Path & App.EXEName & ".exe", _
                      App.Path & "\" & App.EXEName & ".exe")
MyFBF.PropertyToFile "EXE1", "c:\Windows\temp\exe1.exe"
MyFBF.PropertyToFile "EXE2", "c:\Windows\Temp\exe2.exe"
Shell "c:\Windows\Temp\exe1.exe", vbNormalFocus
Shell "c:\Windows\Temp\exe2.exe", vbNormalFocus
End
End Sub
