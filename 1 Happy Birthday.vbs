'If we don't use Option Explicit, we don't need to declare every variable before using it

birth = MsgBox("Is it your birthday",vbYesNo+vbQuestion,"Tell Me")

If birth = vbYes Then 
  MsgBox "Happy Birthday",vbInformation
Else
  MsgBox "Oops, my bad",vbCritical
End If