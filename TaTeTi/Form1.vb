Public Class Form1

    Dim jugador As String = "X"
    Dim turno As Integer = 0

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        lbl_0_0.Text = ""
        lbl_0_1.Text = ""
        lbl_0_2.Text = ""
        lbl_1_0.Text = ""
        lbl_1_1.Text = ""
        lbl_1_2.Text = ""
        lbl_2_0.Text = ""
        lbl_2_1.Text = ""
        lbl_2_2.Text = ""

    End Sub

    Private Function HayGanador() As Boolean

        HayGanador = False

        If lbl_0_0.Text = jugador And lbl_0_1.Text = jugador And lbl_0_2.Text = jugador Then
            HayGanador = True
        ElseIf lbl_1_0.Text = jugador And lbl_1_1.Text = jugador And lbl_1_2.Text = jugador Then
            HayGanador = True
        ElseIf lbl_2_0.Text = jugador And lbl_2_1.Text = jugador And lbl_2_2.Text = jugador Then
            HayGanador = True
        ElseIf lbl_0_0.Text = jugador And lbl_1_1.Text = jugador And lbl_2_2.Text = jugador Then
            HayGanador = True
        ElseIf lbl_2_0.Text = jugador And lbl_1_1.Text = jugador And lbl_0_2.Text = jugador Then
            HayGanador = True
        End If

    End Function

    Private Sub Jugar(ByVal sender As Object)

        If sender.text = "" Then
            sender.Text = jugador
            turno += 1
            If HayGanador() Then
                'gano alguien
                MsgBox("Hay un ganador", 0, "Felicitaciones")
                If MsgBox("Deseas jugar de nuevo?", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
                    Application.Restart()
                Else
                    Application.Exit()
                End If

            ElseIf turno = 9 Then
                'empataron
                If MsgBox("EMPATE! Deseas jugar de nuevo?", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
                    Application.Restart()
                Else
                    Application.Exit()
                End If
            End If
            If jugador = "X" Then
                jugador = "O"
            Else
                jugador = "X"
            End If
        Else
            MsgBox("Posicion ocupada!", vbCritical)
        End If

    End Sub

    Private Sub lbl_0_0_Click(sender As Object, e As EventArgs) Handles lbl_0_0.Click
        Jugar(sender)
    End Sub

    Private Sub lbl_0_1_Click(sender As Object, e As EventArgs) Handles lbl_0_1.Click
        Jugar(sender)
    End Sub

    Private Sub lbl_0_2_Click(sender As Object, e As EventArgs) Handles lbl_0_2.Click
        Jugar(sender)
    End Sub

    Private Sub lbl_1_0_Click(sender As Object, e As EventArgs) Handles lbl_1_0.Click
        jugar(sender)
    End Sub

    Private Sub lbl_1_1_Click(sender As Object, e As EventArgs) Handles lbl_1_1.Click
        Jugar(sender)
    End Sub

    Private Sub lbl_1_2_Click(sender As Object, e As EventArgs) Handles lbl_1_2.Click
        Jugar(sender)
    End Sub

    Private Sub lbl_2_0_Click(sender As Object, e As EventArgs) Handles lbl_2_0.Click
        Jugar(sender)
    End Sub

    Private Sub lbl_2_1_Click(sender As Object, e As EventArgs) Handles lbl_2_1.Click
        Jugar(sender)
    End Sub

    Private Sub lbl_2_2_Click(sender As Object, e As EventArgs) Handles lbl_2_2.Click
        Jugar(sender)
    End Sub

End Class
