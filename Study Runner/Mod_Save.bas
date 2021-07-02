Attribute VB_Name = "Mod_Function"
Option Explicit
Option Base 1
Public Function save()
'세이브파일 저장 함수
On Error GoTo ABC:
Open "C:\Study Runner\Save.SR" For Output As #1
For i = 1 To 12
    Print #1, Stg_Clear(i)
Next
    Print #1, Val(GPoint)
    Print #1, Val(LifPoint)
For i = 1 To 8
    Print #1, Contri(i)
Next

For i = 2 To 12
    Print #1, QuestionB(1, i)
Next

For i = 1 To 8
    Print #1, QuestionB(2, i)
Next

For i = 1 To 9
    Print #1, QuestionB(3, i)
Next
Close #1
Exit Function


ABC:
MkDir "C:\Study Runner\"
Open "C:\Study Runner\Save.SR" For Output As #1
For i = 1 To 12
    Print #1, Stg_Clear(i)
Next
    Print #1, Val(GPoint)
    Print #1, Val(LifPoint)
For i = 1 To 8
    Print #1, Contri(i)
Next
For i = 2 To 12
    Print #1, QuestionB(1, i)
Next

For i = 1 To 8
    Print #1, QuestionB(2, i)
Next

For i = 1 To 9
    Print #1, QuestionB(3, i)
Next
Close #1

End Function

Public Function load(Frm1 As Form, Frm2 As Form)
'세이브파일 로드 함수
On Error GoTo Studying:
Dim Ld(50) As String
Open "C:\Study Runner\Save.SR" For Input As #1
For i = 1 To 50
    Line Input #1, Ld(i)
Next
Close #1
For i = 1 To 12
    If Ld(i) = "True" Then
        Stg_Clear(i) = True
    Else
        Stg_Clear(i) = False
    End If
Next
GPoint = Val(Ld(13))
LifPoint = Val(Ld(14))

For i = 15 To 22
    Contri(i - 14) = Ld(i)
Next

For i = 23 To 33
    QuestionB(1, i - 21) = Ld(i)
Next

For i = 34 To 41
    QuestionB(2, i - 33) = Ld(i)
Next

For i = 42 To 50
    QuestionB(3, i - 41) = Ld(i)
Next

Frm1.Show
Unload Frm2
Exit Function


Studying:
MsgBox "현재 버전에 맞는 정상적인 세이브 파일을 가지고 있지 않습니다."
End Function

Public Function Wrong() As Integer
'정답 여부 판정
MsgBox "틀렸습니다ㅠㅠ..."
Wrong = 1
If item(3) >= 1 Then
    MsgBox "수정테이프 사용 !"
    item(3) = item(3) - 1
Else
    LifPoint = LifPoint - 1
    Wrong = 2
    DoEvents
    If LifPoint <= 0 Then
        MsgBox "Game Over!"
        Wrong = 3
        If MsgBox("랭킹을 전송하시겠습니까?", vbOKCancel) = vbOK Then
            FrmRank.Show
            Unload Frm_Play
            Unload Frm_Ques
        Else
            End
        End If
    End If
End If
End Function

Public Function QuesSel(Num As Integer) As String
'문제 범주 선택
If Num = 1 Then
    QuesSel = "Sci"
ElseIf Num = 2 Then
    QuesSel = "Soc"
Else
    QuesSel = "Non"
End If
End Function
