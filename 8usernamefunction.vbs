Function GetUserName
  'Prompts the user for his name. If the user refuses to provide 
  'his name five times, we give up and return a zero-length string.


  Dim UserName
  Dim TryAgain
  Dim LoopCount


  LoopCount = 1
  Do
    TryAgain = "No"


    UserName = InputBox("Please enter your name:")


    UserName = Trim(UserName)
    If UserName = "" Then
      If LoopCount < 5 Then
        UserName = ""
        TryAgain = "No"
      Else
        MsgBox "You must enter your name."
                               ����<<<����                                 ������                                                          �����������  �����������������������������  ���                    ���                                                          �����������                  �����������������������                                                          ??>