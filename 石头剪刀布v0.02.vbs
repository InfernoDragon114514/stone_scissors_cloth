'initialize handlanguage
Dim gesture(2)
gesture(0) = " ʯͷ "
gesture(1) = " ���� "
gesture(2) = " �� "
'initialize win variable
wins = 0
'initialize random number
Randomize
'show title
MsgBox"STONE&SCISSORS&CLOTH"
'game start
For i=1 To 5
user = CInt(InputBox("ѡ����Ҫ�������ƣ�0: ʯͷ��1: ������2: ���������ʤ��"))
computer = CInt(Rnd * 2)
s = " player��" & gesture(user) & "��COM��" & gesture(computer)
If user = computer Then
MsgBox s & " ƽ�� "
draws = draws + 1
ElseIf computer = (user + 1) Mod 3 Then
MsgBox s & " ��Ӯ�ˣ� "
wins = wins + 1
Else
MsgBox s & " �����ˣ� "
loses = loses + 1
End If
Next
'show score
MsgBox " SCORE�� " & wins
If wins >=3 Then
MsgBox"You win!" 
ElseIf draws>=3 Then
MsgBox"Draw!"
ElseIf loses>=3 Then
MsgBox"Sorry!Try again!" 
Else
MsgBox"Draw!"
End If