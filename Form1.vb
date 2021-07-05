Imports System.IO
Imports System.Text
Imports MSScriptControl
Imports System.Runtime.InteropServices

Public Class Form1
    Public Class SaveClass
        Public killsum As Integer
        Public kd As Single
        Public match As Integer
        Public vic As Integer
        Public dethsum As Integer
    End Class

    Dim num As Integer
    Dim sum As Integer
    Dim kd As Single
    Dim kdst As String
    Dim dethsum As Integer
    Dim deth As Integer
    Dim vic As Integer
    Dim match As Integer

    Private Sub plusBtn_Click(sender As Object, e As EventArgs) Handles plusBtn.Click
        num = CInt(Fix(TextBox2.Text))
        num += 1
        TextBox2.Text = num

    End Sub

    Private Sub minusBtn_Click(sender As Object, e As EventArgs) Handles minusBtn.Click
        num = CInt(Fix(TextBox2.Text))
        If num > 0 Then
            num -= 1
        End If
        TextBox2.Text = num

    End Sub

    Private Sub addBtn_Click(sender As Object, e As EventArgs) Handles addBtn.Click
        num = CInt(Fix(TextBox2.Text))
        sum = CInt(Fix(TextBox1.Text))
        match = CInt(Fix(matchBox1.Text))
        vic = CInt(Fix(victoryBox1.Text))
        deth = CInt(Fix(dethBox1.Text))

        If deth = 0 Then
            vic += 1
        End If

        sum += num
        match += 1
        dethsum += deth

        If dethsum = 0 Then
            kd = sum
        Else
            kd = sum / dethsum
        End If

        kdst = kd.ToString("0.00")
        TextBox1.Text = sum
        matchBox1.Text = match
        victoryBox1.Text = vic
        kideBox1.Text = kdst
    End Sub

    Private Sub plusBtn2_Click(sender As Object, e As EventArgs) Handles plusBtn2.Click
        deth = CInt(Fix(dethBox1.Text))
        deth += 1
        dethBox1.Text = deth

    End Sub

    Private Sub minusBtn2_Click(sender As Object, e As EventArgs) Handles minusBtn2.Click
        deth = CInt(Fix(dethBox1.Text))
        If deth > 0 Then
            deth -= 1
        End If
        dethBox1.Text = deth

    End Sub

    Private Sub SaveItem1_Click(sender As Object, e As EventArgs) Handles SaveItem1.Click
        Dim Ret As DialogResult
        Dim FilePath As String = String.Empty
        Dim obj As New SaveClass()
        obj.killsum = sum
        obj.kd = kdst
        obj.dethsum = dethsum
        obj.match = match
        obj.vic = vic

        Try

            Using Dialog As New SaveFileDialog()

                With Dialog
                    .FileName = "save.xml"
                    .Title = "名前を付けて保存"
                    .Filter = "XMLファイル(*.xml)|*.xml|すべてのファイル(*.*)|*.*"
                    .FilterIndex = 2
                    .RestoreDirectory = True
                End With

                Ret = Dialog.ShowDialog()

                If Ret = DialogResult.OK Then

                    FilePath = Dialog.FileName
                    Dim serializer As New System.Xml.Serialization.XmlSerializer(
                        GetType(SaveClass))
                    Using sw As New StreamWriter(FilePath, False, New System.Text.UTF8Encoding(False))

                        serializer.Serialize(sw, obj)

                    End Using

                End If

            End Using

        Catch ex As Exception

            MessageBox.Show(Err.Description,
                "エラー",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error)

        End Try
    End Sub

    Private Sub LoadItem1_Click(sender As Object, e As EventArgs) Handles LoadItem1.Click
        Dim Ret As DialogResult
        Dim FilePath As String = String.Empty
        Dim SelectFile As String = String.Empty

        Try

            Using Dialog As New OpenFileDialog()

                With Dialog
                    .FileName = "save.xml"
                    .Title = "読み込むデータを選択"
                    .Filter = "XMLファイル(*.xml)|*.xml|すべてのファイル(*.*)|*.*"
                    .FilterIndex = 2
                    .RestoreDirectory = True
                End With

                Ret = Dialog.ShowDialog()

                If Ret = DialogResult.OK Then

                    FilePath = Dialog.FileName
                    Dim serializer As New System.Xml.Serialization.XmlSerializer(
                                           GetType(SaveClass))
                    Using sr As New StreamReader(FilePath, New System.Text.UTF8Encoding(False))
                        Dim obj As SaveClass =
                        DirectCast(serializer.Deserialize(sr), SaveClass)
                        TextBox1.Text = obj.killsum
                        kideBox1.Text = obj.kd
                        dethsum = obj.dethsum
                        matchBox1.Text = obj.match
                        victoryBox1.Text = obj.vic

                    End Using
                End If

            End Using

        Catch ex As Exception

            MessageBox.Show(Err.Description,
                "エラー",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error)

        End Try
    End Sub

    Private Sub TweetItem1_Click(sender As Object, e As EventArgs) Handles TweetItem1.Click

        Dim wsh = CreateObject("WScript.Shell")
        Dim ie = CreateObject("InternetExplorer.Application")
        ie.Visible = True

    End Sub
End Class

