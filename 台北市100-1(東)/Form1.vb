Public Class Form1
    Dim Box(15, 15) As Button
    Dim NowUser As Boolean
    Dim ck As Boolean
    Dim SetNX() As Integer = {-1, -1, -1, 0, 0, 1, 1, 1} 'BOX(X,Y)
    Dim SetNY() As Integer = {-1, 0, 1, -1, 1, -1, 0, 1}
    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        For i As Integer = 1 To 15
            For x As Integer = 1 To 15
                Box(i, x) = Me.Controls("Button" & (i - 1) * 15 + x)
                AddHandler Box(i, x).Click, AddressOf Tread
            Next
        Next
        Re()
    End Sub

    Sub Re() '重新
        NowUser = False : ck = False
        For i As Integer = 1 To 15
            For x As Integer = 1 To 15
                Box(i, x).BackColor = Color.White
                Box(i, x).Text = ""
                Box(i, x).Tag = 0
                Box(i, x).Image = Nothing
            Next
        Next
    End Sub

    Sub Tread(ByVal sender As System.Object, ByVal e As System.EventArgs) '放棋子
        Dim NowX, NowY, BlackFr, WhiteFr As Integer
        Dim Boxck(15, 15) As Boolean : ReDim Boxck(15, 15)
        NowX = (Val(Mid(CType(sender, Button).Name, 7)) - 1) \ 15 + 1
        NowY = (Val(Mid(CType(sender, Button).Name, 7)) - 1) Mod 15 + 1
        If Box(NowX, NowY).Tag = 0 Then
            If NowUser = False Then '白棋
                NowUser = True
                Box(NowX, NowY).Image = New Bitmap("0.png")
                Box(NowX, NowY).Tag = 1
                WhiteFr = 0
                For i As Integer = 0 To 7
                    Track(NowX, NowY, WhiteFr, 1, Boxck, i)
                    For x As Integer = 1 To 15
                        For y As Integer = 1 To 15
                            Boxck(x, y) = False
                        Next
                    Next
                Next
            ElseIf NowUser = True Then '黑棋
                NowUser = False
                Box(NowX, NowY).Image = New Bitmap("1.png")
                Box(NowX, NowY).Tag = 2
                BlackFr = 0
                For i As Integer = 0 To 7
                    Track(NowX, NowY, BlackFr, 2, Boxck, i)
                    For x As Integer = 1 To 15
                        For y As Integer = 1 To 15
                            Boxck(x, y) = False
                        Next
                    Next
                Next
            End If
        End If
    End Sub

    Sub Track(ByVal NowX As Integer, ByVal NowY As Integer, ByVal Fr As Integer, ByVal Now As Integer, ByVal BoxCk(,) As Boolean, ByVal cnt As Integer) '偵測 當前點選X,Y位子 FR=幾個棋子連載一起
        If ck = True Then Exit Sub
        If Fr = 5 Then
            If Now = 1 Then
                MsgBox("白旗贏了") : ck = True : Exit Sub
            ElseIf Now = 2 Then
                MsgBox("黑旗贏了") : ck = True : Exit Sub
            End If
        End If
        If (NowX <= 0 Or NowX >= 16) Or (NowY <= 0 Or NowY >= 16) Then Exit Sub
        If Box(NowX, NowY).Tag = Now And BoxCk(NowX, NowY) = False Then
            BoxCk(NowX, NowY) = True
            Track(NowX + SetNX(cnt), NowY + SetNY(cnt), Fr + 1, Now, BoxCk, cnt)
        End If
    End Sub

    Private Sub 重新開始ToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles 重新開始ToolStripMenuItem.Click
        Re()
    End Sub
End Class
