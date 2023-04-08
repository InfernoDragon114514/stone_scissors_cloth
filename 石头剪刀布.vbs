'initialize handlanguage
Dim gesture(2)
gesture(0) = " 石头 "
gesture(1) = " 剪刀 "
gesture(2) = " 布 "
'initialize win variable
wins = 0
'initialize random number
Randomize
'show title
MsgBox"STONE&SCISSORS&CLOTH"
'game start
For i=1 To 5
user = CInt(InputBox("选择您要出的手势：0: 石头、1: 剪刀、2: 布（五局三胜）"))
computer = CInt(Rnd * 2)
s = " player：" & gesture(user) & "、COM：" & gesture(computer)
If user = computer Then
MsgBox s & " 平局 "
ElseIf computer = (user + 1) Mod 3 Then
MsgBox s & " 您赢了！ "
wins = wins + 1
Else
MsgBox s & " 您输了！ "
End If
Next
'show score
MsgBox " SCORE： " & wins
If wins >=3 Then MsgBox"You win!"
If wins<3 Then MsgBox"Sorry!Try again"