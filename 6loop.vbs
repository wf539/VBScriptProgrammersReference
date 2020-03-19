Dim Greeting
Dim UserName
Dim TryAgain

Do
  
  TryAgain = "No"

  UserName = InputBox("Please enter your name:")
  UserName = Trim(UserName)

  If UserName = "" Then
    MsgBox "You must enter your name."
    TryAgain = "Yes"
  Else
   Greeting = "Hello, " & UserName & ", it's a pleasure to meet you."
End If

Loop While TryAgain = "Yes"

MsgBox Greeting