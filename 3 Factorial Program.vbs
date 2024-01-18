Option Explicit 'Mandates to declare all variables using Dim, before using them

Dim num
num = CInt(InputBox("Enter the Number","Facorial calculator by Saptarshi Dey"))

If num < 0 Then
  Call MsgBox("Negative numbers are not supported",vbCritical,"Error")
Else
  FACTORIAL num
End If

'Another way of calling:- call FACTORIAL(num)

Sub FACTORIAL(x)
  dim a
  dim count
  a = 1
  If x>0 Then 
    For count = 1 to x
      a = a * count
    Next
  End If
  Call MsgBox("The Factorial is of "&x&" is "&a,vbInformation,"Answer is ready")
End Sub