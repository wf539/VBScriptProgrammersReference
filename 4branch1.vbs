Dim Greeting
Dim UserName

UserName = InputBox("Please enter your name:")

UserName = Trim(UserName)

If UserName = "" Then
   Greeting = "Why won't you tell me your name? That's not very nice." 
Else
   Greeting = "Hello, " & UserName & ", it's a pleasure to meet you."
End If

MsgBox Greeting