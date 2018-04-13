Imports System.Data.OleDb
Public Class edit
    Dim pilihan As String



    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Try
            If RadioButton1.Checked = True Then
                pilihan = RadioButton1.Text
            ElseIf RadioButton2.Checked = True Then
                pilihan = RadioButton2.Text
            End If
            Call koneksi()
            Dim edit As String = "update rplb set Nama='" & TextBox2.Text & "', Tanggal_Lahir='" & TextBox3.Text & "', Jenis_Kelamin='" & pilihan & "', Alamat='" & TextBox4.Text & "', No_HP='" & TextBox5.Text & "' where NIS ='" & TextBox1.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call Form1.tampilgrid()
        Catch ex As Exception
        End Try
        Me.Close()
    End Sub

    Private Sub TextBox1_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox1.KeyPress
        On Error Resume Next
        If e.KeyChar = Chr(13) Then
            Call koneksi()
            Call Form1.carikode()
            If dr.HasRows Then
                Call Form1.ketemu()

            End If
        End If
        If Not ((e.KeyChar >= "0" And e.KeyChar <= "9") Or e.KeyChar = vbBack) Then
            e.Handled = True
        End If
    End Sub

    Private Sub TextBox1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox1.TextChanged

    End Sub

    Private Sub edit_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub
End Class