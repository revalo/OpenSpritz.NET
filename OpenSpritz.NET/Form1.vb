Public Class Form1
    Dim bm As Bitmap
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        bm = New Bitmap(PictureBox1.Width, PictureBox1.Height)
        ComboBox1.SelectedIndex = 3
    End Sub
    Sub DrawWord(ByVal word As String, ByVal pivot As Integer)
        Dim fontRender As New Font("Lucida Console", 22)
        Dim center As New Point(bm.Width / 2 - 70, (bm.Height / 2) - TextRenderer.MeasureText("a", fontRender).Height)
        Dim h As Integer = center.Y
        Const LINE_PAD As Integer = 17
        Using g As Graphics = Graphics.FromImage(bm)
            g.Clear(Color.White)

            Dim pen As New Pen(Brushes.Black, 2)

            g.DrawLine(pen, 0, h - LINE_PAD, bm.Width, h - LINE_PAD)
            g.DrawLine(pen, 0, (center.Y + TextRenderer.MeasureText("a", fontRender).Height) + LINE_PAD, bm.Width, (center.Y + TextRenderer.MeasureText("a", fontRender).Height) + LINE_PAD)

            Dim prevChunk As String = word.Substring(0, pivot - 0)
            Dim nextChunk As String = word.Substring(pivot + 1, word.Length - pivot - 1)

            Dim prevLen As Integer = TextRenderer.MeasureText(prevChunk, fontRender).Width
            Dim nextLen As Integer = TextRenderer.MeasureText(nextChunk, fontRender).Width

            Dim pt As Point = New Point(center.X - TextRenderer.MeasureText(word.ToCharArray()(pivot), fontRender).Width, center.Y)
            g.DrawString(word.ToCharArray()(pivot), fontRender, Brushes.Red, pt)

            Const LINE_HEIGHT As Integer = 6
            g.DrawLine(pen, pt.X + Integer.Parse(Math.Round(TextRenderer.MeasureText(word.ToCharArray()(pivot), fontRender).Width / 2, 0)), h - LINE_PAD, pt.X + Integer.Parse(Math.Round(TextRenderer.MeasureText(word.ToCharArray()(pivot), fontRender).Width / 2, 0)), h - LINE_PAD + LINE_HEIGHT)
            g.DrawLine(pen, pt.X + Integer.Parse(Math.Round(TextRenderer.MeasureText(word.ToCharArray()(pivot), fontRender).Width / 2, 0)), (center.Y + TextRenderer.MeasureText("a", fontRender).Height) + LINE_PAD, pt.X + Integer.Parse(Math.Round(TextRenderer.MeasureText(word.ToCharArray()(pivot), fontRender).Width / 2, 0)), (center.Y + TextRenderer.MeasureText("a", fontRender).Height) + LINE_PAD - LINE_HEIGHT)

            Dim prevPt As Point = New Point(pt.X - prevLen + 15, pt.Y)
            g.DrawString(prevChunk, fontRender, Brushes.Black, prevPt)

            Dim nextPt As Point = New Point(pt.X + TextRenderer.MeasureText(word.ToCharArray()(pivot), fontRender).Width - 15, pt.Y)
            g.DrawString(nextChunk, fontRender, Brushes.Black, nextPt)
        End Using

        Me.Invoke(Sub()
                      PictureBox1.Image = bm
                  End Sub)
    End Sub
    Function Slice(ByVal word As String, ByVal st As Integer, ByVal en As Integer) As String
        Return word.Substring(st, en - st)
    End Function
    Function GetPivot(ByVal word As String) As Integer
        Dim bestLetter As Integer = 1
        Dim length As Integer = word.Length

        Select Case length
            Case 1
                bestLetter = 0
            Case 2, 3, 4, 5
                bestLetter = 1
            Case 6, 7, 8, 9
                bestLetter = 2
            Case 10, 11, 12, 13
                bestLetter = 3
            Case Else
                bestLetter = 4
        End Select
        If word.ToCharArray()(bestLetter) = " " Then
            bestLetter = bestLetter - 1
        End If
        Return bestLetter
    End Function
    Private Sub PictureBox1_Click(sender As Object, e As EventArgs) Handles PictureBox1.Click

    End Sub
    Dim txt As String = "This is super awesome, isnt it?"
    Function GetMult(ByVal word As String) As Integer
        Dim multiplier
        Dim wordLength = word.Length
        If wordLength > 5 AndAlso wordLength <= 15 Then
            multiplier = 2 * (wordLength - 5)
        ElseIf wordLength > 15 Then
            multiplier = 25
        Else
            multiplier = 0
        End If
        Return multiplier
    End Function
    Dim pause As Boolean = False
    Dim restart As Boolean = False
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        restart = True
        Threading.Thread.Sleep(1000)
        restart = False
        txt = TextBox1.Text
        Dim t As New Threading.Thread(Sub(wpm As Integer)
                                          'Countdown
                                          DrawWord("3", 0)
                                          Threading.Thread.Sleep(300)
                                          DrawWord("2", 0)
                                          Threading.Thread.Sleep(300)
                                          DrawWord("1", 0)
                                          Threading.Thread.Sleep(300)

                                          Dim words As String() = txt.Split(New [Char]() {" "c, ControlChars.Lf})
                                          Dim sleep As Integer = 0
                                          For Each word In words
                                              If restart Then
                                                  restart = False
                                                  Exit Sub
                                              End If
                                              If Not pause Then
                                                  DrawWord(word, GetPivot(word))
                                                  sleep = (1 + GetMult(word) / 50) * (1 / wpm) * 60000
                                                  If word.EndsWith(",") Then
                                                      sleep *= 1.6
                                                  End If
                                                  If word.EndsWith(".") Or word.EndsWith("?") Or word.EndsWith("!") Then
                                                      sleep *= 1.8
                                                  End If
                                                  If word.EndsWith(":") Then
                                                      sleep *= 1.3
                                                  End If
                                                  Threading.Thread.Sleep(sleep)
                                              Else
                                                  While pause

                                                  End While
                                              End If
                                          Next
                                      End Sub)
        t.Start(ComboBox1.SelectedItem)
    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs)

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        pause = Not pause
        If Button2.Text = "Pause" Then
            Button2.Text = "Play"
        Else
            Button2.Text = "Pause"
        End If
    End Sub
End Class
