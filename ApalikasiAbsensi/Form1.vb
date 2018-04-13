Imports System.Data.OleDb
Imports vbexcel = Microsoft.Office.Interop.Excel

Public Class Form1
    Sub kosong2()
        TextBox3.Text = ""
        TextBox4.Text = ""
        TextBox5.Text = ""
        RadioButton1.Checked = False
        RadioButton2.Checked = False
        RadioButton3.Checked = False
        RadioButton4.Checked = False
        Label9.Text = "-"
        Label11.Text = "-"
        Label13.Text = "-"
        Label15.Text = "-"


        TextBox3.Focus()
    End Sub
    Sub carikode2()
        If DateTimePicker1.Text = "01 Juli 2016" Or DateTimePicker1.Text = "02 Juli 2016" Or DateTimePicker1.Text = "03 Juli 2016" Or DateTimePicker1.Text = "04 Juli 2016" Or DateTimePicker1.Text = "05 Juli 2016" Or DateTimePicker1.Text = "06 Juli 2016" Or DateTimePicker1.Text = "07 Juli 2016" Or DateTimePicker1.Text = "08 Juli 2016" Or DateTimePicker1.Text = "09 Juli 2016" Or DateTimePicker1.Text = "10 Juli 2016" Or DateTimePicker1.Text = "11 Juli 2016" Or DateTimePicker1.Text = "12 Juli 2016" Or DateTimePicker1.Text = "13 Juli 2016" Or DateTimePicker1.Text = "14 Juli 2016" Or DateTimePicker1.Text = "15 Juli 2016" Or DateTimePicker1.Text = "16 Juli 2016" Or DateTimePicker1.Text = "17 Juli 2016" Or DateTimePicker1.Text = "18 Juli 2016" Or DateTimePicker1.Text = "19 Juli 2016" Or DateTimePicker1.Text = "20 Juli 2016" Or DateTimePicker1.Text = "21 Juli 2016" Or DateTimePicker1.Text = "22 Juli 2016" Or DateTimePicker1.Text = "23 Juli 2016" Or DateTimePicker1.Text = "24 Juli 2016" Or DateTimePicker1.Text = "25 Juli 2016" Or DateTimePicker1.Text = "26 Juli 2016" Or DateTimePicker1.Text = "27 Juli 2016" Or DateTimePicker1.Text = "28 Juli 2016" Or DateTimePicker1.Text = "29 Juli 2016" Or DateTimePicker1.Text = "30 Juli 2016" Or DateTimePicker1.Text = "31 Juli 2016" Then
            cmd = New OleDbCommand("select * from rplb_juli where NIS = '" & TextBox3.Text & "'", conn)
            dr = cmd.ExecuteReader
            dr.Read()
        ElseIf DateTimePicker1.Text = "01 Agustus 2016" Or DateTimePicker1.Text = "02 Agustus 2016" Or DateTimePicker1.Text = "03 Agustus 2016" Or DateTimePicker1.Text = "04 Agustus 2016" Or DateTimePicker1.Text = "05 Agustus 2016" Or DateTimePicker1.Text = "06 Agustus 2016" Or DateTimePicker1.Text = "07 Agustus 2016" Or DateTimePicker1.Text = "08 Agustus 2016" Or DateTimePicker1.Text = "09 Agustus 2016" Or DateTimePicker1.Text = "10 Agustus 2016" Or DateTimePicker1.Text = "11 Agustus 2016" Or DateTimePicker1.Text = "12 Agustus 2016" Or DateTimePicker1.Text = "13 Agustus 2016" Or DateTimePicker1.Text = "14 Agustus 2016" Or DateTimePicker1.Text = "15 Agustus 2016" Or DateTimePicker1.Text = "16 Agustus 2016" Or DateTimePicker1.Text = "17 Agustus 2016" Or DateTimePicker1.Text = "18 Agustus 2016" Or DateTimePicker1.Text = "19 Agustus 2016" Or DateTimePicker1.Text = "20 Agustus 2016" Or DateTimePicker1.Text = "21 Agustus 2016" Or DateTimePicker1.Text = "22 Agustus 2016" Or DateTimePicker1.Text = "23 Agustus 2016" Or DateTimePicker1.Text = "24 Agustus 2016" Or DateTimePicker1.Text = "25 Agustus 2016" Or DateTimePicker1.Text = "26 Agustus 2016" Or DateTimePicker1.Text = "27 Agustus 2016" Or DateTimePicker1.Text = "28 Agustus 2016" Or DateTimePicker1.Text = "29 Agustus 2016" Or DateTimePicker1.Text = "30 Agustus 2016" Or DateTimePicker1.Text = "31 Agustus 2016" Then
            cmd = New OleDbCommand("select * from rplb_agustus where NIS = '" & TextBox3.Text & "'", conn)
            dr = cmd.ExecuteReader
            dr.Read()
        ElseIf DateTimePicker1.Text = "01 September 2016" Or DateTimePicker1.Text = "02 September 2016" Or DateTimePicker1.Text = "03 September 2016" Or DateTimePicker1.Text = "04 September 2016" Or DateTimePicker1.Text = "05 September 2016" Or DateTimePicker1.Text = "06 September 2016" Or DateTimePicker1.Text = "07 September 2016" Or DateTimePicker1.Text = "08 September 2016" Or DateTimePicker1.Text = "09 September 2016" Or DateTimePicker1.Text = "10 September 2016" Or DateTimePicker1.Text = "11 September 2016" Or DateTimePicker1.Text = "12 September 2016" Or DateTimePicker1.Text = "13 September 2016" Or DateTimePicker1.Text = "14 September 2016" Or DateTimePicker1.Text = "15 September 2016" Or DateTimePicker1.Text = "16 September 2016" Or DateTimePicker1.Text = "17 September 2016" Or DateTimePicker1.Text = "18 September 2016" Or DateTimePicker1.Text = "19 September 2016" Or DateTimePicker1.Text = "20 September 2016" Or DateTimePicker1.Text = "21 September 2016" Or DateTimePicker1.Text = "22 September 2016" Or DateTimePicker1.Text = "23 September 2016" Or DateTimePicker1.Text = "24 September 2016" Or DateTimePicker1.Text = "25 September 2016" Or DateTimePicker1.Text = "26 September 2016" Or DateTimePicker1.Text = "27 September 2016" Or DateTimePicker1.Text = "28 September 2016" Or DateTimePicker1.Text = "29 September 2016" Or DateTimePicker1.Text = "30 September 2016" Then
            cmd = New OleDbCommand("select * from rplb_september where NIS = '" & TextBox3.Text & "'", conn)
            dr = cmd.ExecuteReader
            dr.Read()
        ElseIf DateTimePicker1.Text = "01 Oktober 2016" Or DateTimePicker1.Text = "02 Oktober 2016" Or DateTimePicker1.Text = "03 Oktober 2016" Or DateTimePicker1.Text = "04 Oktober 2016" Or DateTimePicker1.Text = "05 Oktober 2016" Or DateTimePicker1.Text = "06 Oktober 2016" Or DateTimePicker1.Text = "07 Oktober 2016" Or DateTimePicker1.Text = "08 Oktober 2016" Or DateTimePicker1.Text = "09 Oktober 2016" Or DateTimePicker1.Text = "10 Oktober 2016" Or DateTimePicker1.Text = "11 Oktober 2016" Or DateTimePicker1.Text = "12 Oktober 2016" Or DateTimePicker1.Text = "13 Oktober 2016" Or DateTimePicker1.Text = "14 Oktober 2016" Or DateTimePicker1.Text = "15 Oktober 2016" Or DateTimePicker1.Text = "16 Oktober 2016" Or DateTimePicker1.Text = "17 Oktober 2016" Or DateTimePicker1.Text = "18 Oktober 2016" Or DateTimePicker1.Text = "19 Oktober 2016" Or DateTimePicker1.Text = "20 Oktober 2016" Or DateTimePicker1.Text = "21 Oktober 2016" Or DateTimePicker1.Text = "22 Oktober 2016" Or DateTimePicker1.Text = "23 Oktober 2016" Or DateTimePicker1.Text = "24 Oktober 2016" Or DateTimePicker1.Text = "25 Oktober 2016" Or DateTimePicker1.Text = "26 Oktober 2016" Or DateTimePicker1.Text = "27 Oktober 2016" Or DateTimePicker1.Text = "28 Oktober 2016" Or DateTimePicker1.Text = "29 Oktober 2016" Or DateTimePicker1.Text = "30 Oktober 2016" Or DateTimePicker1.Text = "31 Oktober 2016" Then
            cmd = New OleDbCommand("select * from rplb_oktober where NIS = '" & TextBox3.Text & "'", conn)
            dr = cmd.ExecuteReader
            dr.Read()
        ElseIf DateTimePicker1.Text = "01 November 2016" Or DateTimePicker1.Text = "02 November 2016" Or DateTimePicker1.Text = "03 November 2016" Or DateTimePicker1.Text = "04 November 2016" Or DateTimePicker1.Text = "05 November 2016" Or DateTimePicker1.Text = "06 November 2016" Or DateTimePicker1.Text = "07 November 2016" Or DateTimePicker1.Text = "08 November 2016" Or DateTimePicker1.Text = "09 November 2016" Or DateTimePicker1.Text = "10 November 2016" Or DateTimePicker1.Text = "11 November 2016" Or DateTimePicker1.Text = "12 November 2016" Or DateTimePicker1.Text = "13 November 2016" Or DateTimePicker1.Text = "14 November 2016" Or DateTimePicker1.Text = "15 November 2016" Or DateTimePicker1.Text = "16 November 2016" Or DateTimePicker1.Text = "17 November 2016" Or DateTimePicker1.Text = "18 November 2016" Or DateTimePicker1.Text = "19 November 2016" Or DateTimePicker1.Text = "20 November 2016" Or DateTimePicker1.Text = "21 November 2016" Or DateTimePicker1.Text = "22 November 2016" Or DateTimePicker1.Text = "23 November 2016" Or DateTimePicker1.Text = "24 November 2016" Or DateTimePicker1.Text = "25 November 2016" Or DateTimePicker1.Text = "26 November 2016" Or DateTimePicker1.Text = "27 November 2016" Or DateTimePicker1.Text = "28 November 2016" Or DateTimePicker1.Text = "29 November 2016" Or DateTimePicker1.Text = "30 November 2016" Then
            cmd = New OleDbCommand("select * from rplb_november where NIS = '" & TextBox3.Text & "'", conn)
            dr = cmd.ExecuteReader
            dr.Read()
        ElseIf DateTimePicker1.Text = "01 Desember 2016" Or DateTimePicker1.Text = "02 Desember 2016" Or DateTimePicker1.Text = "03 Desember 2016" Or DateTimePicker1.Text = "04 Desember 2016" Or DateTimePicker1.Text = "05 Desember 2016" Or DateTimePicker1.Text = "06 Desember 2016" Or DateTimePicker1.Text = "07 Desember 2016" Or DateTimePicker1.Text = "08 Desember 2016" Or DateTimePicker1.Text = "09 Desember 2016" Or DateTimePicker1.Text = "10 Desember 2016" Or DateTimePicker1.Text = "11 Desember 2016" Or DateTimePicker1.Text = "12 Desember 2016" Or DateTimePicker1.Text = "13 Desember 2016" Or DateTimePicker1.Text = "14 Desember 2016" Or DateTimePicker1.Text = "15 Desember 2016" Or DateTimePicker1.Text = "16 Desember 2016" Or DateTimePicker1.Text = "17 Desember 2016" Or DateTimePicker1.Text = "18 Desember 2016" Or DateTimePicker1.Text = "19 Desember 2016" Or DateTimePicker1.Text = "20 Desember 2016" Or DateTimePicker1.Text = "21 Desember 2016" Or DateTimePicker1.Text = "22 Desember 2016" Or DateTimePicker1.Text = "23 Desember 2016" Or DateTimePicker1.Text = "24 Desember 2016" Or DateTimePicker1.Text = "25 Desember 2016" Or DateTimePicker1.Text = "26 Desember 2016" Or DateTimePicker1.Text = "27 Desember 2016" Or DateTimePicker1.Text = "28 Desember 2016" Or DateTimePicker1.Text = "29 Desember 2016" Or DateTimePicker1.Text = "30 Desember 2016" Or DateTimePicker1.Text = "31 Desember 2016" Then
            cmd = New OleDbCommand("select * from rplb_desember where NIS = '" & TextBox3.Text & "'", conn)
            dr = cmd.ExecuteReader
            dr.Read()
        ElseIf DateTimePicker1.Text = "01 Januari 2017" Or DateTimePicker1.Text = "02 Januari 2017" Or DateTimePicker1.Text = "03 Januari 2017" Or DateTimePicker1.Text = "04 Januari 2017" Or DateTimePicker1.Text = "05 Januari 2017" Or DateTimePicker1.Text = "06 Januari 2017" Or DateTimePicker1.Text = "07 Januari 2017" Or DateTimePicker1.Text = "08 Januari 2017" Or DateTimePicker1.Text = "09 Januari 2017" Or DateTimePicker1.Text = "10 Januari 2017" Or DateTimePicker1.Text = "11 Januari 2017" Or DateTimePicker1.Text = "12 Januari 2017" Or DateTimePicker1.Text = "13 Januari 2017" Or DateTimePicker1.Text = "14 Januari 2017" Or DateTimePicker1.Text = "15 Januari 2017" Or DateTimePicker1.Text = "16 Januari 2017" Or DateTimePicker1.Text = "17 Januari 2017" Or DateTimePicker1.Text = "18 Januari 2017" Or DateTimePicker1.Text = "19 Januari 2017" Or DateTimePicker1.Text = "20 Januari 2017" Or DateTimePicker1.Text = "21 Januari 2017" Or DateTimePicker1.Text = "22 Januari 2017" Or DateTimePicker1.Text = "23 Januari 2017" Or DateTimePicker1.Text = "24 Januari 2017" Or DateTimePicker1.Text = "25 Januari 2017" Or DateTimePicker1.Text = "26 Januari 2017" Or DateTimePicker1.Text = "27 Januari 2017" Or DateTimePicker1.Text = "28 Januari 2017" Or DateTimePicker1.Text = "29 Januari 2017" Or DateTimePicker1.Text = "30 Januari 2017" Or DateTimePicker1.Text = "31 Januari 2017" Then
            cmd = New OleDbCommand("select * from rplb_januari where NIS = '" & TextBox3.Text & "'", conn)
            dr = cmd.ExecuteReader
            dr.Read()
        ElseIf DateTimePicker1.Text = "01 Februari 2017" Or DateTimePicker1.Text = "02 Februari 2017" Or DateTimePicker1.Text = "03 Februari 2017" Or DateTimePicker1.Text = "04 Februari 2017" Or DateTimePicker1.Text = "05 Februari 2017" Or DateTimePicker1.Text = "06 Februari 2017" Or DateTimePicker1.Text = "07 Februari 2017" Or DateTimePicker1.Text = "08 Februari 2017" Or DateTimePicker1.Text = "09 Februari 2017" Or DateTimePicker1.Text = "10 Februari 2017" Or DateTimePicker1.Text = "11 Februari 2017" Or DateTimePicker1.Text = "12 Februari 2017" Or DateTimePicker1.Text = "13 Februari 2017" Or DateTimePicker1.Text = "14 Februari 2017" Or DateTimePicker1.Text = "15 Februari 2017" Or DateTimePicker1.Text = "16 Februari 2017" Or DateTimePicker1.Text = "17 Februari 2017" Or DateTimePicker1.Text = "18 Februari 2017" Or DateTimePicker1.Text = "19 Februari 2017" Or DateTimePicker1.Text = "20 Februari 2017" Or DateTimePicker1.Text = "21 Februari 2017" Or DateTimePicker1.Text = "22 Februari 2017" Or DateTimePicker1.Text = "23 Februari 2017" Or DateTimePicker1.Text = "24 Februari 2017" Or DateTimePicker1.Text = "25 Februari 2017" Or DateTimePicker1.Text = "26 Februari 2017" Or DateTimePicker1.Text = "27 Februari 2017" Or DateTimePicker1.Text = "28 Februari 2017" Then
            cmd = New OleDbCommand("select * from rplb_februari where NIS = '" & TextBox3.Text & "'", conn)
            dr = cmd.ExecuteReader
            dr.Read()
        ElseIf DateTimePicker1.Text = "01 Maret 2017" Or DateTimePicker1.Text = "02 Maret 2017" Or DateTimePicker1.Text = "03 Maret 2017" Or DateTimePicker1.Text = "04 Maret 2017" Or DateTimePicker1.Text = "05 Maret 2017" Or DateTimePicker1.Text = "06 Maret 2017" Or DateTimePicker1.Text = "07 Maret 2017" Or DateTimePicker1.Text = "08 Maret 2017" Or DateTimePicker1.Text = "09 Maret 2017" Or DateTimePicker1.Text = "10 Maret 2017" Or DateTimePicker1.Text = "11 Maret 2017" Or DateTimePicker1.Text = "12 Maret 2017" Or DateTimePicker1.Text = "13 Maret 2017" Or DateTimePicker1.Text = "14 Maret 2017" Or DateTimePicker1.Text = "15 Maret 2017" Or DateTimePicker1.Text = "16 Maret 2017" Or DateTimePicker1.Text = "17 Maret 2017" Or DateTimePicker1.Text = "18 Maret 2017" Or DateTimePicker1.Text = "19 Maret 2017" Or DateTimePicker1.Text = "20 Maret 2017" Or DateTimePicker1.Text = "21 Maret 2017" Or DateTimePicker1.Text = "22 Maret 2017" Or DateTimePicker1.Text = "23 Maret 2017" Or DateTimePicker1.Text = "24 Maret 2017" Or DateTimePicker1.Text = "25 Maret 2017" Or DateTimePicker1.Text = "26 Maret 2017" Or DateTimePicker1.Text = "27 Maret 2017" Or DateTimePicker1.Text = "28 Maret 2017" Or DateTimePicker1.Text = "29 Maret 2017" Or DateTimePicker1.Text = "30 Maret 2017" Or DateTimePicker1.Text = "31 Maret 2017" Then
            cmd = New OleDbCommand("select * from rplb_maret where NIS = '" & TextBox3.Text & "'", conn)
            dr = cmd.ExecuteReader
            dr.Read()
        ElseIf DateTimePicker1.Text = "01 April 2017" Or DateTimePicker1.Text = "02 April 2017" Or DateTimePicker1.Text = "03 April 2017" Or DateTimePicker1.Text = "04 April 2017" Or DateTimePicker1.Text = "05 April 2017" Or DateTimePicker1.Text = "06 April 2017" Or DateTimePicker1.Text = "07 April 2017" Or DateTimePicker1.Text = "08 April 2017" Or DateTimePicker1.Text = "09 April 2017" Or DateTimePicker1.Text = "10 April 2017" Or DateTimePicker1.Text = "11 April 2017" Or DateTimePicker1.Text = "12 April 2017" Or DateTimePicker1.Text = "13 April 2017" Or DateTimePicker1.Text = "14 April 2017" Or DateTimePicker1.Text = "15 April 2017" Or DateTimePicker1.Text = "16 April 2017" Or DateTimePicker1.Text = "17 April 2017" Or DateTimePicker1.Text = "18 April 2017" Or DateTimePicker1.Text = "19 April 2017" Or DateTimePicker1.Text = "20 April 2017" Or DateTimePicker1.Text = "21 April 2017" Or DateTimePicker1.Text = "22 April 2017" Or DateTimePicker1.Text = "23 April 2017" Or DateTimePicker1.Text = "24 April 2017" Or DateTimePicker1.Text = "25 April 2017" Or DateTimePicker1.Text = "26 April 2017" Or DateTimePicker1.Text = "27 April 2017" Or DateTimePicker1.Text = "28 April 2017" Or DateTimePicker1.Text = "29 April 2017" Or DateTimePicker1.Text = "30 April 2017" Then
            cmd = New OleDbCommand("select * from rplb_april where NIS = '" & TextBox3.Text & "'", conn)
            dr = cmd.ExecuteReader
            dr.Read()
        ElseIf DateTimePicker1.Text = "01 Mei 2017" Or DateTimePicker1.Text = "02 Mei 2017" Or DateTimePicker1.Text = "03 Mei 2017" Or DateTimePicker1.Text = "04 Mei 2017" Or DateTimePicker1.Text = "05 Mei 2017" Or DateTimePicker1.Text = "06 Mei 2017" Or DateTimePicker1.Text = "07 Mei 2017" Or DateTimePicker1.Text = "08 Mei 2017" Or DateTimePicker1.Text = "09 Mei 2017" Or DateTimePicker1.Text = "10 Mei 2017" Or DateTimePicker1.Text = "11 Mei 2017" Or DateTimePicker1.Text = "12 Mei 2017" Or DateTimePicker1.Text = "13 Mei 2017" Or DateTimePicker1.Text = "14 Mei 2017" Or DateTimePicker1.Text = "15 Mei 2017" Or DateTimePicker1.Text = "16 Mei 2017" Or DateTimePicker1.Text = "17 Mei 2017" Or DateTimePicker1.Text = "18 Mei 2017" Or DateTimePicker1.Text = "19 Mei 2017" Or DateTimePicker1.Text = "20 Mei 2017" Or DateTimePicker1.Text = "21 Mei 2017" Or DateTimePicker1.Text = "22 Mei 2017" Or DateTimePicker1.Text = "23 Mei 2017" Or DateTimePicker1.Text = "24 Mei 2017" Or DateTimePicker1.Text = "25 Mei 2017" Or DateTimePicker1.Text = "26 Mei 2017" Or DateTimePicker1.Text = "27 Mei 2017" Or DateTimePicker1.Text = "28 Mei 2017" Or DateTimePicker1.Text = "29 Mei 2017" Or DateTimePicker1.Text = "30 Mei 2017" Or DateTimePicker1.Text = "31 Mei 2017" Then
            cmd = New OleDbCommand("select * from rplb_mei where NIS = '" & TextBox3.Text & "'", conn)
            dr = cmd.ExecuteReader
            dr.Read()
        ElseIf DateTimePicker1.Text = "01 Juni 2017" Or DateTimePicker1.Text = "02 Juni 2017" Or DateTimePicker1.Text = "03 Juni 2017" Or DateTimePicker1.Text = "04 Juni 2017" Or DateTimePicker1.Text = "05 Juni 2017" Or DateTimePicker1.Text = "06 Juni 2017" Or DateTimePicker1.Text = "07 Juni 2017" Or DateTimePicker1.Text = "08 Juni 2017" Or DateTimePicker1.Text = "09 Juni 2017" Or DateTimePicker1.Text = "10 Juni 2017" Or DateTimePicker1.Text = "11 Juni 2017" Or DateTimePicker1.Text = "12 Juni 2017" Or DateTimePicker1.Text = "13 Juni 2017" Or DateTimePicker1.Text = "14 Juni 2017" Or DateTimePicker1.Text = "15 Juni 2017" Or DateTimePicker1.Text = "16 Juni 2017" Or DateTimePicker1.Text = "17 Juni 2017" Or DateTimePicker1.Text = "18 Juni 2017" Or DateTimePicker1.Text = "19 Juni 2017" Or DateTimePicker1.Text = "20 Juni 2017" Or DateTimePicker1.Text = "21 Juni 2017" Or DateTimePicker1.Text = "22 Juni 2017" Or DateTimePicker1.Text = "23 Juni 2017" Or DateTimePicker1.Text = "24 Juni 2017" Or DateTimePicker1.Text = "25 Juni 2017" Or DateTimePicker1.Text = "26 Juni 2017" Or DateTimePicker1.Text = "27 Juni 2017" Or DateTimePicker1.Text = "28 Juni 2017" Or DateTimePicker1.Text = "29 Juni 2017" Or DateTimePicker1.Text = "30 Juni 2017" Then
            cmd = New OleDbCommand("select * from rplb_juni where NIS = '" & TextBox3.Text & "'", conn)
            dr = cmd.ExecuteReader
            dr.Read()
        End If
    End Sub

    Sub ketemu2()
        TextBox3.Text = dr.Item(0)
        TextBox4.Text = dr.Item(1)
        TextBox5.Text = dr.Item(33)

    End Sub

    Sub jumlah()
        Dim total, total2 As Integer
        total = dgv.Rows.Count - 1
        Label2.Text = total
        total2 = dgv.Rows.Count - 1
        Label3.Text = total2


    End Sub

    Sub tampilgrid()
        da = New OleDbDataAdapter("select * from rplb", conn)
        ds = New DataSet
        da.Fill(ds)
        dgv.DataSource = ds.Tables(0)

    End Sub

    Sub carikode()
        cmd = New OleDbCommand("select * from rplb where NIS = '" & edit.TextBox1.Text & "'", conn)
        dr = cmd.ExecuteReader
        dr.Read()
    End Sub

    Sub ketemu()
        edit.TextBox1.Text = dr.Item(0)
        edit.TextBox2.Text = dr.Item(1)
        edit.TextBox3.Text = dr.Item(2)
        edit.TextBox4.Text = dr.Item(4)
        edit.TextBox5.Text = dr.Item(5)
        If dr.Item(3) = "Laki-laki" Then
            edit.RadioButton1.Checked = True
        Else
            If dr.Item(3) = "Perempuan" Then
                edit.RadioButton2.Checked = True
            End If
        End If
    End Sub

    Dim pilihan As String
    Sub kosong()

  

    End Sub


    Sub tampilgrid1()

        If DateTimePicker1.Text = "01 Juli 2016" Or DateTimePicker1.Text = "02 Juli 2016" Or DateTimePicker1.Text = "03 Juli 2016" Or DateTimePicker1.Text = "04 Juli 2016" Or DateTimePicker1.Text = "05 Juli 2016" Or DateTimePicker1.Text = "06 Juli 2016" Or DateTimePicker1.Text = "07 Juli 2016" Or DateTimePicker1.Text = "08 Juli 2016" Or DateTimePicker1.Text = "09 Juli 2016" Or DateTimePicker1.Text = "10 Juli 2016" Or DateTimePicker1.Text = "11 Juli 2016" Or DateTimePicker1.Text = "12 Juli 2016" Or DateTimePicker1.Text = "13 Juli 2016" Or DateTimePicker1.Text = "14 Juli 2016" Or DateTimePicker1.Text = "15 Juli 2016" Or DateTimePicker1.Text = "16 Juli 2016" Or DateTimePicker1.Text = "17 Juli 2016" Or DateTimePicker1.Text = "18 Juli 2016" Or DateTimePicker1.Text = "19 Juli 2016" Or DateTimePicker1.Text = "20 Juli 2016" Or DateTimePicker1.Text = "21 Juli 2016" Or DateTimePicker1.Text = "22 Juli 2016" Or DateTimePicker1.Text = "23 Juli 2016" Or DateTimePicker1.Text = "24 Juli 2016" Or DateTimePicker1.Text = "25 Juli 2016" Or DateTimePicker1.Text = "26 Juli 2016" Or DateTimePicker1.Text = "27 Juli 2016" Or DateTimePicker1.Text = "28 Juli 2016" Or DateTimePicker1.Text = "29 Juli 2016" Or DateTimePicker1.Text = "30 Juli 2016" Or DateTimePicker1.Text = "31 Juli 2016" Then
            da = New OleDbDataAdapter("select * from rplb_juli", conn)
            ds = New DataSet
            da.Fill(ds)
            dgv2.DataSource = ds.Tables(0)
        ElseIf DateTimePicker1.Text = "01 Agustus 2016" Or DateTimePicker1.Text = "02 Agustus 2016" Or DateTimePicker1.Text = "03 Agustus 2016" Or DateTimePicker1.Text = "04 Agustus 2016" Or DateTimePicker1.Text = "05 Agustus 2016" Or DateTimePicker1.Text = "06 Agustus 2016" Or DateTimePicker1.Text = "07 Agustus 2016" Or DateTimePicker1.Text = "08 Agustus 2016" Or DateTimePicker1.Text = "09 Agustus 2016" Or DateTimePicker1.Text = "10 Agustus 2016" Or DateTimePicker1.Text = "11 Agustus 2016" Or DateTimePicker1.Text = "12 Agustus 2016" Or DateTimePicker1.Text = "13 Agustus 2016" Or DateTimePicker1.Text = "14 Agustus 2016" Or DateTimePicker1.Text = "15 Agustus 2016" Or DateTimePicker1.Text = "16 Agustus 2016" Or DateTimePicker1.Text = "17 Agustus 2016" Or DateTimePicker1.Text = "18 Agustus 2016" Or DateTimePicker1.Text = "19 Agustus 2016" Or DateTimePicker1.Text = "20 Agustus 2016" Or DateTimePicker1.Text = "21 Agustus 2016" Or DateTimePicker1.Text = "22 Agustus 2016" Or DateTimePicker1.Text = "23 Agustus 2016" Or DateTimePicker1.Text = "24 Agustus 2016" Or DateTimePicker1.Text = "25 Agustus 2016" Or DateTimePicker1.Text = "26 Agustus 2016" Or DateTimePicker1.Text = "27 Agustus 2016" Or DateTimePicker1.Text = "28 Agustus 2016" Or DateTimePicker1.Text = "29 Agustus 2016" Or DateTimePicker1.Text = "30 Agustus 2016" Or DateTimePicker1.Text = "31 Agustus 2016" Then
            da = New OleDbDataAdapter("select * from rplb_agustus", conn)
            ds = New DataSet
            da.Fill(ds)
            dgv2.DataSource = ds.Tables(0)
        ElseIf DateTimePicker1.Text = "01 September 2016" Or DateTimePicker1.Text = "02 September 2016" Or DateTimePicker1.Text = "03 September 2016" Or DateTimePicker1.Text = "04 September 2016" Or DateTimePicker1.Text = "05 September 2016" Or DateTimePicker1.Text = "06 September 2016" Or DateTimePicker1.Text = "07 September 2016" Or DateTimePicker1.Text = "08 September 2016" Or DateTimePicker1.Text = "09 September 2016" Or DateTimePicker1.Text = "10 September 2016" Or DateTimePicker1.Text = "11 September 2016" Or DateTimePicker1.Text = "12 September 2016" Or DateTimePicker1.Text = "13 September 2016" Or DateTimePicker1.Text = "14 September 2016" Or DateTimePicker1.Text = "15 September 2016" Or DateTimePicker1.Text = "16 September 2016" Or DateTimePicker1.Text = "17 September 2016" Or DateTimePicker1.Text = "18 September 2016" Or DateTimePicker1.Text = "19 September 2016" Or DateTimePicker1.Text = "20 September 2016" Or DateTimePicker1.Text = "21 September 2016" Or DateTimePicker1.Text = "22 September 2016" Or DateTimePicker1.Text = "23 September 2016" Or DateTimePicker1.Text = "24 September 2016" Or DateTimePicker1.Text = "25 September 2016" Or DateTimePicker1.Text = "26 September 2016" Or DateTimePicker1.Text = "27 September 2016" Or DateTimePicker1.Text = "28 September 2016" Or DateTimePicker1.Text = "29 September 2016" Or DateTimePicker1.Text = "30 September 2016" Then
            da = New OleDbDataAdapter("select * from rplb_september", conn)
            ds = New DataSet
            da.Fill(ds)
            dgv2.DataSource = ds.Tables(0)
        ElseIf DateTimePicker1.Text = "01 Oktober 2016" Or DateTimePicker1.Text = "02 Oktober 2016" Or DateTimePicker1.Text = "03 Oktober 2016" Or DateTimePicker1.Text = "04 Oktober 2016" Or DateTimePicker1.Text = "05 Oktober 2016" Or DateTimePicker1.Text = "06 Oktober 2016" Or DateTimePicker1.Text = "07 Oktober 2016" Or DateTimePicker1.Text = "08 Oktober 2016" Or DateTimePicker1.Text = "09 Oktober 2016" Or DateTimePicker1.Text = "10 Oktober 2016" Or DateTimePicker1.Text = "11 Oktober 2016" Or DateTimePicker1.Text = "12 Oktober 2016" Or DateTimePicker1.Text = "13 Oktober 2016" Or DateTimePicker1.Text = "14 Oktober 2016" Or DateTimePicker1.Text = "15 Oktober 2016" Or DateTimePicker1.Text = "16 Oktober 2016" Or DateTimePicker1.Text = "17 Oktober 2016" Or DateTimePicker1.Text = "18 Oktober 2016" Or DateTimePicker1.Text = "19 Oktober 2016" Or DateTimePicker1.Text = "20 Oktober 2016" Or DateTimePicker1.Text = "21 Oktober 2016" Or DateTimePicker1.Text = "22 Oktober 2016" Or DateTimePicker1.Text = "23 Oktober 2016" Or DateTimePicker1.Text = "24 Oktober 2016" Or DateTimePicker1.Text = "25 Oktober 2016" Or DateTimePicker1.Text = "26 Oktober 2016" Or DateTimePicker1.Text = "27 Oktober 2016" Or DateTimePicker1.Text = "28 Oktober 2016" Or DateTimePicker1.Text = "29 Oktober 2016" Or DateTimePicker1.Text = "30 Oktober 2016" Or DateTimePicker1.Text = "31 Oktober 2016" Then
            da = New OleDbDataAdapter("select * from rplb_oktober", conn)
            ds = New DataSet
            da.Fill(ds)
            dgv2.DataSource = ds.Tables(0)
        ElseIf DateTimePicker1.Text = "01 November 2016" Or DateTimePicker1.Text = "02 November 2016" Or DateTimePicker1.Text = "03 November 2016" Or DateTimePicker1.Text = "04 November 2016" Or DateTimePicker1.Text = "05 November 2016" Or DateTimePicker1.Text = "06 November 2016" Or DateTimePicker1.Text = "07 November 2016" Or DateTimePicker1.Text = "08 November 2016" Or DateTimePicker1.Text = "09 November 2016" Or DateTimePicker1.Text = "10 November 2016" Or DateTimePicker1.Text = "11 November 2016" Or DateTimePicker1.Text = "12 November 2016" Or DateTimePicker1.Text = "13 November 2016" Or DateTimePicker1.Text = "14 November 2016" Or DateTimePicker1.Text = "15 November 2016" Or DateTimePicker1.Text = "16 November 2016" Or DateTimePicker1.Text = "17 November 2016" Or DateTimePicker1.Text = "18 November 2016" Or DateTimePicker1.Text = "19 November 2016" Or DateTimePicker1.Text = "20 November 2016" Or DateTimePicker1.Text = "21 November 2016" Or DateTimePicker1.Text = "22 November 2016" Or DateTimePicker1.Text = "23 November 2016" Or DateTimePicker1.Text = "24 November 2016" Or DateTimePicker1.Text = "25 November 2016" Or DateTimePicker1.Text = "26 November 2016" Or DateTimePicker1.Text = "27 November 2016" Or DateTimePicker1.Text = "28 November 2016" Or DateTimePicker1.Text = "29 November 2016" Or DateTimePicker1.Text = "30 November 2016" Then
            da = New OleDbDataAdapter("select * from rplb_november", conn)
            ds = New DataSet
            da.Fill(ds)
            dgv2.DataSource = ds.Tables(0)
        ElseIf DateTimePicker1.Text = "01 Desember 2016" Or DateTimePicker1.Text = "02 Desember 2016" Or DateTimePicker1.Text = "03 Desember 2016" Or DateTimePicker1.Text = "04 Desember 2016" Or DateTimePicker1.Text = "05 Desember 2016" Or DateTimePicker1.Text = "06 Desember 2016" Or DateTimePicker1.Text = "07 Desember 2016" Or DateTimePicker1.Text = "08 Desember 2016" Or DateTimePicker1.Text = "09 Desember 2016" Or DateTimePicker1.Text = "10 Desember 2016" Or DateTimePicker1.Text = "11 Desember 2016" Or DateTimePicker1.Text = "12 Desember 2016" Or DateTimePicker1.Text = "13 Desember 2016" Or DateTimePicker1.Text = "14 Desember 2016" Or DateTimePicker1.Text = "15 Desember 2016" Or DateTimePicker1.Text = "16 Desember 2016" Or DateTimePicker1.Text = "17 Desember 2016" Or DateTimePicker1.Text = "18 Desember 2016" Or DateTimePicker1.Text = "19 Desember 2016" Or DateTimePicker1.Text = "20 Desember 2016" Or DateTimePicker1.Text = "21 Desember 2016" Or DateTimePicker1.Text = "22 Desember 2016" Or DateTimePicker1.Text = "23 Desember 2016" Or DateTimePicker1.Text = "24 Desember 2016" Or DateTimePicker1.Text = "25 Desember 2016" Or DateTimePicker1.Text = "26 Desember 2016" Or DateTimePicker1.Text = "27 Desember 2016" Or DateTimePicker1.Text = "28 Desember 2016" Or DateTimePicker1.Text = "29 Desember 2016" Or DateTimePicker1.Text = "30 Desember 2016" Or DateTimePicker1.Text = "31 Desember 2016" Then
            da = New OleDbDataAdapter("select * from rplb_desember", conn)
            ds = New DataSet
            da.Fill(ds)
            dgv2.DataSource = ds.Tables(0)
        ElseIf DateTimePicker1.Text = "01 Januari 2017" Or DateTimePicker1.Text = "02 Januari 2017" Or DateTimePicker1.Text = "03 Januari 2017" Or DateTimePicker1.Text = "04 Januari 2017" Or DateTimePicker1.Text = "05 Januari 2017" Or DateTimePicker1.Text = "06 Januari 2017" Or DateTimePicker1.Text = "07 Januari 2017" Or DateTimePicker1.Text = "08 Januari 2017" Or DateTimePicker1.Text = "09 Januari 2017" Or DateTimePicker1.Text = "10 Januari 2017" Or DateTimePicker1.Text = "11 Januari 2017" Or DateTimePicker1.Text = "12 Januari 2017" Or DateTimePicker1.Text = "13 Januari 2017" Or DateTimePicker1.Text = "14 Januari 2017" Or DateTimePicker1.Text = "15 Januari 2017" Or DateTimePicker1.Text = "16 Januari 2017" Or DateTimePicker1.Text = "17 Januari 2017" Or DateTimePicker1.Text = "18 Januari 2017" Or DateTimePicker1.Text = "19 Januari 2017" Or DateTimePicker1.Text = "20 Januari 2017" Or DateTimePicker1.Text = "21 Januari 2017" Or DateTimePicker1.Text = "22 Januari 2017" Or DateTimePicker1.Text = "23 Januari 2017" Or DateTimePicker1.Text = "24 Januari 2017" Or DateTimePicker1.Text = "25 Januari 2017" Or DateTimePicker1.Text = "26 Januari 2017" Or DateTimePicker1.Text = "27 Januari 2017" Or DateTimePicker1.Text = "28 Januari 2017" Or DateTimePicker1.Text = "29 Januari 2017" Or DateTimePicker1.Text = "30 Januari 2017" Or DateTimePicker1.Text = "31 Januari 2017" Then
            da = New OleDbDataAdapter("select * from rplb_januari", conn)
            ds = New DataSet
            da.Fill(ds)
            dgv2.DataSource = ds.Tables(0)
        ElseIf DateTimePicker1.Text = "01 Februari 2017" Or DateTimePicker1.Text = "02 Februari 2017" Or DateTimePicker1.Text = "03 Februari 2017" Or DateTimePicker1.Text = "04 Februari 2017" Or DateTimePicker1.Text = "05 Februari 2017" Or DateTimePicker1.Text = "06 Februari 2017" Or DateTimePicker1.Text = "07 Februari 2017" Or DateTimePicker1.Text = "08 Februari 2017" Or DateTimePicker1.Text = "09 Februari 2017" Or DateTimePicker1.Text = "10 Februari 2017" Or DateTimePicker1.Text = "11 Februari 2017" Or DateTimePicker1.Text = "12 Februari 2017" Or DateTimePicker1.Text = "13 Februari 2017" Or DateTimePicker1.Text = "14 Februari 2017" Or DateTimePicker1.Text = "15 Februari 2017" Or DateTimePicker1.Text = "16 Februari 2017" Or DateTimePicker1.Text = "17 Februari 2017" Or DateTimePicker1.Text = "18 Februari 2017" Or DateTimePicker1.Text = "19 Februari 2017" Or DateTimePicker1.Text = "20 Februari 2017" Or DateTimePicker1.Text = "21 Februari 2017" Or DateTimePicker1.Text = "22 Februari 2017" Or DateTimePicker1.Text = "23 Februari 2017" Or DateTimePicker1.Text = "24 Februari 2017" Or DateTimePicker1.Text = "25 Februari 2017" Or DateTimePicker1.Text = "26 Februari 2017" Or DateTimePicker1.Text = "27 Februari 2017" Or DateTimePicker1.Text = "28 Februari 2017" Then
            da = New OleDbDataAdapter("select * from rplb_februari", conn)
            ds = New DataSet
            da.Fill(ds)
            dgv2.DataSource = ds.Tables(0)
        ElseIf DateTimePicker1.Text = "01 Maret 2017" Or DateTimePicker1.Text = "02 Maret 2017" Or DateTimePicker1.Text = "03 Maret 2017" Or DateTimePicker1.Text = "04 Maret 2017" Or DateTimePicker1.Text = "05 Maret 2017" Or DateTimePicker1.Text = "06 Maret 2017" Or DateTimePicker1.Text = "07 Maret 2017" Or DateTimePicker1.Text = "08 Maret 2017" Or DateTimePicker1.Text = "09 Maret 2017" Or DateTimePicker1.Text = "10 Maret 2017" Or DateTimePicker1.Text = "11 Maret 2017" Or DateTimePicker1.Text = "12 Maret 2017" Or DateTimePicker1.Text = "13 Maret 2017" Or DateTimePicker1.Text = "14 Maret 2017" Or DateTimePicker1.Text = "15 Maret 2017" Or DateTimePicker1.Text = "16 Maret 2017" Or DateTimePicker1.Text = "17 Maret 2017" Or DateTimePicker1.Text = "18 Maret 2017" Or DateTimePicker1.Text = "19 Maret 2017" Or DateTimePicker1.Text = "20 Maret 2017" Or DateTimePicker1.Text = "21 Maret 2017" Or DateTimePicker1.Text = "22 Maret 2017" Or DateTimePicker1.Text = "23 Maret 2017" Or DateTimePicker1.Text = "24 Maret 2017" Or DateTimePicker1.Text = "25 Maret 2017" Or DateTimePicker1.Text = "26 Maret 2017" Or DateTimePicker1.Text = "27 Maret 2017" Or DateTimePicker1.Text = "28 Maret 2017" Or DateTimePicker1.Text = "29 Maret 2017" Or DateTimePicker1.Text = "30 Maret 2017" Or DateTimePicker1.Text = "31 Maret 2017" Then
            da = New OleDbDataAdapter("select * from rplb_maret", conn)
            ds = New DataSet
            da.Fill(ds)
            dgv2.DataSource = ds.Tables(0)
        ElseIf DateTimePicker1.Text = "01 April 2017" Or DateTimePicker1.Text = "02 April 2017" Or DateTimePicker1.Text = "03 April 2017" Or DateTimePicker1.Text = "04 April 2017" Or DateTimePicker1.Text = "05 April 2017" Or DateTimePicker1.Text = "06 April 2017" Or DateTimePicker1.Text = "07 April 2017" Or DateTimePicker1.Text = "08 April 2017" Or DateTimePicker1.Text = "09 April 2017" Or DateTimePicker1.Text = "10 April 2017" Or DateTimePicker1.Text = "11 April 2017" Or DateTimePicker1.Text = "12 April 2017" Or DateTimePicker1.Text = "13 April 2017" Or DateTimePicker1.Text = "14 April 2017" Or DateTimePicker1.Text = "15 April 2017" Or DateTimePicker1.Text = "16 April 2017" Or DateTimePicker1.Text = "17 April 2017" Or DateTimePicker1.Text = "18 April 2017" Or DateTimePicker1.Text = "19 April 2017" Or DateTimePicker1.Text = "20 April 2017" Or DateTimePicker1.Text = "21 April 2017" Or DateTimePicker1.Text = "22 April 2017" Or DateTimePicker1.Text = "23 April 2017" Or DateTimePicker1.Text = "24 April 2017" Or DateTimePicker1.Text = "25 April 2017" Or DateTimePicker1.Text = "26 April 2017" Or DateTimePicker1.Text = "27 April 2017" Or DateTimePicker1.Text = "28 April 2017" Or DateTimePicker1.Text = "29 April 2017" Or DateTimePicker1.Text = "30 April 2017" Then
            da = New OleDbDataAdapter("select * from rplb_april", conn)
            ds = New DataSet
            da.Fill(ds)
            dgv2.DataSource = ds.Tables(0)
        ElseIf DateTimePicker1.Text = "01 Mei 2017" Or DateTimePicker1.Text = "02 Mei 2017" Or DateTimePicker1.Text = "03 Mei 2017" Or DateTimePicker1.Text = "04 Mei 2017" Or DateTimePicker1.Text = "05 Mei 2017" Or DateTimePicker1.Text = "06 Mei 2017" Or DateTimePicker1.Text = "07 Mei 2017" Or DateTimePicker1.Text = "08 Mei 2017" Or DateTimePicker1.Text = "09 Mei 2017" Or DateTimePicker1.Text = "10 Mei 2017" Or DateTimePicker1.Text = "11 Mei 2017" Or DateTimePicker1.Text = "12 Mei 2017" Or DateTimePicker1.Text = "13 Mei 2017" Or DateTimePicker1.Text = "14 Mei 2017" Or DateTimePicker1.Text = "15 Mei 2017" Or DateTimePicker1.Text = "16 Mei 2017" Or DateTimePicker1.Text = "17 Mei 2017" Or DateTimePicker1.Text = "18 Mei 2017" Or DateTimePicker1.Text = "19 Mei 2017" Or DateTimePicker1.Text = "20 Mei 2017" Or DateTimePicker1.Text = "21 Mei 2017" Or DateTimePicker1.Text = "22 Mei 2017" Or DateTimePicker1.Text = "23 Mei 2017" Or DateTimePicker1.Text = "24 Mei 2017" Or DateTimePicker1.Text = "25 Mei 2017" Or DateTimePicker1.Text = "26 Mei 2017" Or DateTimePicker1.Text = "27 Mei 2017" Or DateTimePicker1.Text = "28 Mei 2017" Or DateTimePicker1.Text = "29 Mei 2017" Or DateTimePicker1.Text = "30 Mei 2017" Or DateTimePicker1.Text = "31 Mei 2017" Then
            da = New OleDbDataAdapter("select * from rplb_mei", conn)
            ds = New DataSet
            da.Fill(ds)
            dgv2.DataSource = ds.Tables(0)
        ElseIf DateTimePicker1.Text = "01 Juni 2017" Or DateTimePicker1.Text = "02 Juni 2017" Or DateTimePicker1.Text = "03 Juni 2017" Or DateTimePicker1.Text = "04 Juni 2017" Or DateTimePicker1.Text = "05 Juni 2017" Or DateTimePicker1.Text = "06 Juni 2017" Or DateTimePicker1.Text = "07 Juni 2017" Or DateTimePicker1.Text = "08 Juni 2017" Or DateTimePicker1.Text = "09 Juni 2017" Or DateTimePicker1.Text = "10 Juni 2017" Or DateTimePicker1.Text = "11 Juni 2017" Or DateTimePicker1.Text = "12 Juni 2017" Or DateTimePicker1.Text = "13 Juni 2017" Or DateTimePicker1.Text = "14 Juni 2017" Or DateTimePicker1.Text = "15 Juni 2017" Or DateTimePicker1.Text = "16 Juni 2017" Or DateTimePicker1.Text = "17 Juni 2017" Or DateTimePicker1.Text = "18 Juni 2017" Or DateTimePicker1.Text = "19 Juni 2017" Or DateTimePicker1.Text = "20 Juni 2017" Or DateTimePicker1.Text = "21 Juni 2017" Or DateTimePicker1.Text = "22 Juni 2017" Or DateTimePicker1.Text = "23 Juni 2017" Or DateTimePicker1.Text = "24 Juni 2017" Or DateTimePicker1.Text = "25 Juni 2017" Or DateTimePicker1.Text = "26 Juni 2017" Or DateTimePicker1.Text = "27 Juni 2017" Or DateTimePicker1.Text = "28 Juni 2017" Or DateTimePicker1.Text = "29 Juni 2017" Or DateTimePicker1.Text = "30 Juni 2017" Then
            da = New OleDbDataAdapter("select * from rplb_juni", conn)
            ds = New DataSet
            da.Fill(ds)
            dgv2.DataSource = ds.Tables(0)

        End If
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Panel1.Show()
        Panel2.Hide()
    End Sub

    Private Sub Form1_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        splas.Close()
    End Sub

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Call koneksi()
        Call tampilgrid()
        Call jumlah()
        SaveFileDialog1.FileName = ""
        SaveFileDialog1.Filter = "Excel 2003 (*.xls) | *.xls |Excel 2007 (*.xlsx) | *.xlsx"
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        tambah.Show()
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        edit.Show()
    End Sub

    Private Sub DataGridView1_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs)

    End Sub

    Private Sub DataGridView1_CellMouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellMouseEventArgs)
        On Error Resume Next
        edit.TextBox1.Text = dgv.Rows(e.RowIndex).Cells(0).Value
        Call carikode2()
        If dr.HasRows Then
            Call ketemu2()
        End If
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        If edit.TextBox1.Text = "" Then
            MsgBox("Data ada yang kosong")
            Exit Sub
        End If
        Call carikode()
        If Not dr.HasRows Then
            MsgBox("NIS yang anda masukkan tidak terdaftar")
            Exit Sub
        End If
        If MessageBox.Show("Yakin ingin hapus data ini...??", "", MessageBoxButtons.YesNo) = Windows.Forms.DialogResult.Yes Then
            Call koneksi()
            Dim hapus As String = "delete from rplb where NIS= '" & edit.TextBox1.Text & "'"
            cmd = New OleDbCommand(hapus, conn)
            cmd.ExecuteNonQuery()

            Call tampilgrid()
        End If
    End Sub

    Private Sub Label2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub TextBox2_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        Call tampilgrid1()
        Panel2.Show()
        Panel1.Hide()
        Call tampilgrid1()
    End Sub

    Private Sub DateTimePicker1_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)

        If DateTimePicker1.Text = "01 Juli 2016" Or DateTimePicker1.Text = "02 Juli 2016" Or DateTimePicker1.Text = "03 Juli 2016" Or DateTimePicker1.Text = "04 Juli 2016" Or DateTimePicker1.Text = "05 Juli 2016" Or DateTimePicker1.Text = "06 Juli 2016" Or DateTimePicker1.Text = "07 Juli 2016" Or DateTimePicker1.Text = "08 Juli 2016" Or DateTimePicker1.Text = "09 Juli 2016" Or DateTimePicker1.Text = "10 Juli 2016" Or DateTimePicker1.Text = "11 Juli 2016" Or DateTimePicker1.Text = "12 Juli 2016" Or DateTimePicker1.Text = "13 Juli 2016" Or DateTimePicker1.Text = "14 Juli 2016" Or DateTimePicker1.Text = "15 Juli 2016" Or DateTimePicker1.Text = "16 Juli 2016" Or DateTimePicker1.Text = "17 Juli 2016" Or DateTimePicker1.Text = "18 Juli 2016" Or DateTimePicker1.Text = "19 Juli 2016" Or DateTimePicker1.Text = "20 Juli 2016" Or DateTimePicker1.Text = "21 Juli 2016" Or DateTimePicker1.Text = "22 Juli 2016" Or DateTimePicker1.Text = "23 Juli 2016" Or DateTimePicker1.Text = "24 Juli 2016" Or DateTimePicker1.Text = "25 Juli 2016" Or DateTimePicker1.Text = "26 Juli 2016" Or DateTimePicker1.Text = "27 Juli 2016" Or DateTimePicker1.Text = "28 Juli 2016" Or DateTimePicker1.Text = "29 Juli 2016" Or DateTimePicker1.Text = "30 Juli 2016" Or DateTimePicker1.Text = "31 Juli 2016" Then
            da = New OleDbDataAdapter("select * from rplb_juli", conn)
            ds = New DataSet
            da.Fill(ds)
            dgv2.DataSource = ds.Tables(0)
        ElseIf DateTimePicker1.Text = "01 Agustus 2016" Or DateTimePicker1.Text = "02 Agustus 2016" Or DateTimePicker1.Text = "03 Agustus 2016" Or DateTimePicker1.Text = "04 Agustus 2016" Or DateTimePicker1.Text = "05 Agustus 2016" Or DateTimePicker1.Text = "06 Agustus 2016" Or DateTimePicker1.Text = "07 Agustus 2016" Or DateTimePicker1.Text = "08 Agustus 2016" Or DateTimePicker1.Text = "09 Agustus 2016" Or DateTimePicker1.Text = "10 Agustus 2016" Or DateTimePicker1.Text = "11 Agustus 2016" Or DateTimePicker1.Text = "12 Agustus 2016" Or DateTimePicker1.Text = "13 Agustus 2016" Or DateTimePicker1.Text = "14 Agustus 2016" Or DateTimePicker1.Text = "15 Agustus 2016" Or DateTimePicker1.Text = "16 Agustus 2016" Or DateTimePicker1.Text = "17 Agustus 2016" Or DateTimePicker1.Text = "18 Agustus 2016" Or DateTimePicker1.Text = "19 Agustus 2016" Or DateTimePicker1.Text = "20 Agustus 2016" Or DateTimePicker1.Text = "21 Agustus 2016" Or DateTimePicker1.Text = "22 Agustus 2016" Or DateTimePicker1.Text = "23 Agustus 2016" Or DateTimePicker1.Text = "24 Agustus 2016" Or DateTimePicker1.Text = "25 Agustus 2016" Or DateTimePicker1.Text = "26 Agustus 2016" Or DateTimePicker1.Text = "27 Agustus 2016" Or DateTimePicker1.Text = "28 Agustus 2016" Or DateTimePicker1.Text = "29 Agustus 2016" Or DateTimePicker1.Text = "30 Agustus 2016" Or DateTimePicker1.Text = "31 Agustus 2016" Then
            da = New OleDbDataAdapter("select * from rplb_agustus", conn)
            ds = New DataSet
            da.Fill(ds)
            dgv2.DataSource = ds.Tables(0)
        ElseIf DateTimePicker1.Text = "01 September 2016" Or DateTimePicker1.Text = "02 September 2016" Or DateTimePicker1.Text = "03 September 2016" Or DateTimePicker1.Text = "04 September 2016" Or DateTimePicker1.Text = "05 September 2016" Or DateTimePicker1.Text = "06 September 2016" Or DateTimePicker1.Text = "07 September 2016" Or DateTimePicker1.Text = "08 September 2016" Or DateTimePicker1.Text = "09 September 2016" Or DateTimePicker1.Text = "10 September 2016" Or DateTimePicker1.Text = "11 September 2016" Or DateTimePicker1.Text = "12 September 2016" Or DateTimePicker1.Text = "13 September 2016" Or DateTimePicker1.Text = "14 September 2016" Or DateTimePicker1.Text = "15 September 2016" Or DateTimePicker1.Text = "16 September 2016" Or DateTimePicker1.Text = "17 September 2016" Or DateTimePicker1.Text = "18 September 2016" Or DateTimePicker1.Text = "19 September 2016" Or DateTimePicker1.Text = "20 September 2016" Or DateTimePicker1.Text = "21 September 2016" Or DateTimePicker1.Text = "22 September 2016" Or DateTimePicker1.Text = "23 September 2016" Or DateTimePicker1.Text = "24 September 2016" Or DateTimePicker1.Text = "25 September 2016" Or DateTimePicker1.Text = "26 September 2016" Or DateTimePicker1.Text = "27 September 2016" Or DateTimePicker1.Text = "28 September 2016" Or DateTimePicker1.Text = "29 September 2016" Or DateTimePicker1.Text = "30 September 2016" Then
            da = New OleDbDataAdapter("select * from rplb_september", conn)
            ds = New DataSet
            da.Fill(ds)
            dgv2.DataSource = ds.Tables(0)
        ElseIf DateTimePicker1.Text = "01 Oktober 2016" Or DateTimePicker1.Text = "02 Oktober 2016" Or DateTimePicker1.Text = "03 Oktober 2016" Or DateTimePicker1.Text = "04 Oktober 2016" Or DateTimePicker1.Text = "05 Oktober 2016" Or DateTimePicker1.Text = "06 Oktober 2016" Or DateTimePicker1.Text = "07 Oktober 2016" Or DateTimePicker1.Text = "08 Oktober 2016" Or DateTimePicker1.Text = "09 Oktober 2016" Or DateTimePicker1.Text = "10 Oktober 2016" Or DateTimePicker1.Text = "11 Oktober 2016" Or DateTimePicker1.Text = "12 Oktober 2016" Or DateTimePicker1.Text = "13 Oktober 2016" Or DateTimePicker1.Text = "14 Oktober 2016" Or DateTimePicker1.Text = "15 Oktober 2016" Or DateTimePicker1.Text = "16 Oktober 2016" Or DateTimePicker1.Text = "17 Oktober 2016" Or DateTimePicker1.Text = "18 Oktober 2016" Or DateTimePicker1.Text = "19 Oktober 2016" Or DateTimePicker1.Text = "20 Oktober 2016" Or DateTimePicker1.Text = "21 Oktober 2016" Or DateTimePicker1.Text = "22 Oktober 2016" Or DateTimePicker1.Text = "23 Oktober 2016" Or DateTimePicker1.Text = "24 Oktober 2016" Or DateTimePicker1.Text = "25 Oktober 2016" Or DateTimePicker1.Text = "26 Oktober 2016" Or DateTimePicker1.Text = "27 Oktober 2016" Or DateTimePicker1.Text = "28 Oktober 2016" Or DateTimePicker1.Text = "29 Oktober 2016" Or DateTimePicker1.Text = "30 Oktober 2016" Or DateTimePicker1.Text = "31 Oktober 2016" Then
            da = New OleDbDataAdapter("select * from rplb_oktober", conn)
            ds = New DataSet
            da.Fill(ds)
            dgv2.DataSource = ds.Tables(0)
        ElseIf DateTimePicker1.Text = "01 November 2016" Or DateTimePicker1.Text = "02 November 2016" Or DateTimePicker1.Text = "03 November 2016" Or DateTimePicker1.Text = "04 November 2016" Or DateTimePicker1.Text = "05 November 2016" Or DateTimePicker1.Text = "06 November 2016" Or DateTimePicker1.Text = "07 November 2016" Or DateTimePicker1.Text = "08 November 2016" Or DateTimePicker1.Text = "09 November 2016" Or DateTimePicker1.Text = "10 November 2016" Or DateTimePicker1.Text = "11 November 2016" Or DateTimePicker1.Text = "12 November 2016" Or DateTimePicker1.Text = "13 November 2016" Or DateTimePicker1.Text = "14 November 2016" Or DateTimePicker1.Text = "15 November 2016" Or DateTimePicker1.Text = "16 November 2016" Or DateTimePicker1.Text = "17 November 2016" Or DateTimePicker1.Text = "18 November 2016" Or DateTimePicker1.Text = "19 November 2016" Or DateTimePicker1.Text = "20 November 2016" Or DateTimePicker1.Text = "21 November 2016" Or DateTimePicker1.Text = "22 November 2016" Or DateTimePicker1.Text = "23 November 2016" Or DateTimePicker1.Text = "24 November 2016" Or DateTimePicker1.Text = "25 November 2016" Or DateTimePicker1.Text = "26 November 2016" Or DateTimePicker1.Text = "27 November 2016" Or DateTimePicker1.Text = "28 November 2016" Or DateTimePicker1.Text = "29 November 2016" Or DateTimePicker1.Text = "30 November 2016" Then
            da = New OleDbDataAdapter("select * from rplb_november", conn)
            ds = New DataSet
            da.Fill(ds)
            dgv2.DataSource = ds.Tables(0)
        ElseIf DateTimePicker1.Text = "01 Desember 2016" Or DateTimePicker1.Text = "02 Desember 2016" Or DateTimePicker1.Text = "03 Desember 2016" Or DateTimePicker1.Text = "04 Desember 2016" Or DateTimePicker1.Text = "05 Desember 2016" Or DateTimePicker1.Text = "06 Desember 2016" Or DateTimePicker1.Text = "07 Desember 2016" Or DateTimePicker1.Text = "08 Desember 2016" Or DateTimePicker1.Text = "09 Desember 2016" Or DateTimePicker1.Text = "10 Desember 2016" Or DateTimePicker1.Text = "11 Desember 2016" Or DateTimePicker1.Text = "12 Desember 2016" Or DateTimePicker1.Text = "13 Desember 2016" Or DateTimePicker1.Text = "14 Desember 2016" Or DateTimePicker1.Text = "15 Desember 2016" Or DateTimePicker1.Text = "16 Desember 2016" Or DateTimePicker1.Text = "17 Desember 2016" Or DateTimePicker1.Text = "18 Desember 2016" Or DateTimePicker1.Text = "19 Desember 2016" Or DateTimePicker1.Text = "20 Desember 2016" Or DateTimePicker1.Text = "21 Desember 2016" Or DateTimePicker1.Text = "22 Desember 2016" Or DateTimePicker1.Text = "23 Desember 2016" Or DateTimePicker1.Text = "24 Desember 2016" Or DateTimePicker1.Text = "25 Desember 2016" Or DateTimePicker1.Text = "26 Desember 2016" Or DateTimePicker1.Text = "27 Desember 2016" Or DateTimePicker1.Text = "28 Desember 2016" Or DateTimePicker1.Text = "29 Desember 2016" Or DateTimePicker1.Text = "30 Desember 2016" Or DateTimePicker1.Text = "31 Desember 2016" Then
            da = New OleDbDataAdapter("select * from rplb_desember", conn)
            ds = New DataSet
            da.Fill(ds)
            dgv2.DataSource = ds.Tables(0)
        ElseIf DateTimePicker1.Text = "01 Januari 2017" Or DateTimePicker1.Text = "02 Januari 2017" Or DateTimePicker1.Text = "03 Januari 2017" Or DateTimePicker1.Text = "04 Januari 2017" Or DateTimePicker1.Text = "05 Januari 2017" Or DateTimePicker1.Text = "06 Januari 2017" Or DateTimePicker1.Text = "07 Januari 2017" Or DateTimePicker1.Text = "08 Januari 2017" Or DateTimePicker1.Text = "09 Januari 2017" Or DateTimePicker1.Text = "10 Januari 2017" Or DateTimePicker1.Text = "11 Januari 2017" Or DateTimePicker1.Text = "12 Januari 2017" Or DateTimePicker1.Text = "13 Januari 2017" Or DateTimePicker1.Text = "14 Januari 2017" Or DateTimePicker1.Text = "15 Januari 2017" Or DateTimePicker1.Text = "16 Januari 2017" Or DateTimePicker1.Text = "17 Januari 2017" Or DateTimePicker1.Text = "18 Januari 2017" Or DateTimePicker1.Text = "19 Januari 2017" Or DateTimePicker1.Text = "20 Januari 2017" Or DateTimePicker1.Text = "21 Januari 2017" Or DateTimePicker1.Text = "22 Januari 2017" Or DateTimePicker1.Text = "23 Januari 2017" Or DateTimePicker1.Text = "24 Januari 2017" Or DateTimePicker1.Text = "25 Januari 2017" Or DateTimePicker1.Text = "26 Januari 2017" Or DateTimePicker1.Text = "27 Januari 2017" Or DateTimePicker1.Text = "28 Januari 2017" Or DateTimePicker1.Text = "29 Januari 2017" Or DateTimePicker1.Text = "30 Januari 2017" Or DateTimePicker1.Text = "31 Januari 2017" Then
            da = New OleDbDataAdapter("select * from rplb_januari", conn)
            ds = New DataSet
            da.Fill(ds)
            dgv2.DataSource = ds.Tables(0)
        ElseIf DateTimePicker1.Text = "01 Februari 2017" Or DateTimePicker1.Text = "02 Februari 2017" Or DateTimePicker1.Text = "03 Februari 2017" Or DateTimePicker1.Text = "04 Februari 2017" Or DateTimePicker1.Text = "05 Februari 2017" Or DateTimePicker1.Text = "06 Februari 2017" Or DateTimePicker1.Text = "07 Februari 2017" Or DateTimePicker1.Text = "08 Februari 2017" Or DateTimePicker1.Text = "09 Februari 2017" Or DateTimePicker1.Text = "10 Februari 2017" Or DateTimePicker1.Text = "11 Februari 2017" Or DateTimePicker1.Text = "12 Februari 2017" Or DateTimePicker1.Text = "13 Februari 2017" Or DateTimePicker1.Text = "14 Februari 2017" Or DateTimePicker1.Text = "15 Februari 2017" Or DateTimePicker1.Text = "16 Februari 2017" Or DateTimePicker1.Text = "17 Februari 2017" Or DateTimePicker1.Text = "18 Februari 2017" Or DateTimePicker1.Text = "19 Februari 2017" Or DateTimePicker1.Text = "20 Februari 2017" Or DateTimePicker1.Text = "21 Februari 2017" Or DateTimePicker1.Text = "22 Februari 2017" Or DateTimePicker1.Text = "23 Februari 2017" Or DateTimePicker1.Text = "24 Februari 2017" Or DateTimePicker1.Text = "25 Februari 2017" Or DateTimePicker1.Text = "26 Februari 2017" Or DateTimePicker1.Text = "27 Februari 2017" Or DateTimePicker1.Text = "28 Februari 2017" Then
            da = New OleDbDataAdapter("select * from rplb_februari", conn)
            ds = New DataSet
            da.Fill(ds)
            dgv2.DataSource = ds.Tables(0)
        ElseIf DateTimePicker1.Text = "01 Maret 2017" Or DateTimePicker1.Text = "02 Maret 2017" Or DateTimePicker1.Text = "03 Maret 2017" Or DateTimePicker1.Text = "04 Maret 2017" Or DateTimePicker1.Text = "05 Maret 2017" Or DateTimePicker1.Text = "06 Maret 2017" Or DateTimePicker1.Text = "07 Maret 2017" Or DateTimePicker1.Text = "08 Maret 2017" Or DateTimePicker1.Text = "09 Maret 2017" Or DateTimePicker1.Text = "10 Maret 2017" Or DateTimePicker1.Text = "11 Maret 2017" Or DateTimePicker1.Text = "12 Maret 2017" Or DateTimePicker1.Text = "13 Maret 2017" Or DateTimePicker1.Text = "14 Maret 2017" Or DateTimePicker1.Text = "15 Maret 2017" Or DateTimePicker1.Text = "16 Maret 2017" Or DateTimePicker1.Text = "17 Maret 2017" Or DateTimePicker1.Text = "18 Maret 2017" Or DateTimePicker1.Text = "19 Maret 2017" Or DateTimePicker1.Text = "20 Maret 2017" Or DateTimePicker1.Text = "21 Maret 2017" Or DateTimePicker1.Text = "22 Maret 2017" Or DateTimePicker1.Text = "23 Maret 2017" Or DateTimePicker1.Text = "24 Maret 2017" Or DateTimePicker1.Text = "25 Maret 2017" Or DateTimePicker1.Text = "26 Maret 2017" Or DateTimePicker1.Text = "27 Maret 2017" Or DateTimePicker1.Text = "28 Maret 2017" Or DateTimePicker1.Text = "29 Maret 2017" Or DateTimePicker1.Text = "30 Maret 2017" Or DateTimePicker1.Text = "31 Maret 2017" Then
            da = New OleDbDataAdapter("select * from rplb_maret", conn)
            ds = New DataSet
            da.Fill(ds)
            dgv2.DataSource = ds.Tables(0)
        ElseIf DateTimePicker1.Text = "01 April 2017" Or DateTimePicker1.Text = "02 April 2017" Or DateTimePicker1.Text = "03 April 2017" Or DateTimePicker1.Text = "04 April 2017" Or DateTimePicker1.Text = "05 April 2017" Or DateTimePicker1.Text = "06 April 2017" Or DateTimePicker1.Text = "07 April 2017" Or DateTimePicker1.Text = "08 April 2017" Or DateTimePicker1.Text = "09 April 2017" Or DateTimePicker1.Text = "10 April 2017" Or DateTimePicker1.Text = "11 April 2017" Or DateTimePicker1.Text = "12 April 2017" Or DateTimePicker1.Text = "13 April 2017" Or DateTimePicker1.Text = "14 April 2017" Or DateTimePicker1.Text = "15 April 2017" Or DateTimePicker1.Text = "16 April 2017" Or DateTimePicker1.Text = "17 April 2017" Or DateTimePicker1.Text = "18 April 2017" Or DateTimePicker1.Text = "19 April 2017" Or DateTimePicker1.Text = "20 April 2017" Or DateTimePicker1.Text = "21 April 2017" Or DateTimePicker1.Text = "22 April 2017" Or DateTimePicker1.Text = "23 April 2017" Or DateTimePicker1.Text = "24 April 2017" Or DateTimePicker1.Text = "25 April 2017" Or DateTimePicker1.Text = "26 April 2017" Or DateTimePicker1.Text = "27 April 2017" Or DateTimePicker1.Text = "28 April 2017" Or DateTimePicker1.Text = "29 April 2017" Or DateTimePicker1.Text = "30 April 2017" Then
            da = New OleDbDataAdapter("select * from rplb_april", conn)
            ds = New DataSet
            da.Fill(ds)
            dgv2.DataSource = ds.Tables(0)
        ElseIf DateTimePicker1.Text = "01 Mei 2017" Or DateTimePicker1.Text = "02 Mei 2017" Or DateTimePicker1.Text = "03 Mei 2017" Or DateTimePicker1.Text = "04 Mei 2017" Or DateTimePicker1.Text = "05 Mei 2017" Or DateTimePicker1.Text = "06 Mei 2017" Or DateTimePicker1.Text = "07 Mei 2017" Or DateTimePicker1.Text = "08 Mei 2017" Or DateTimePicker1.Text = "09 Mei 2017" Or DateTimePicker1.Text = "10 Mei 2017" Or DateTimePicker1.Text = "11 Mei 2017" Or DateTimePicker1.Text = "12 Mei 2017" Or DateTimePicker1.Text = "13 Mei 2017" Or DateTimePicker1.Text = "14 Mei 2017" Or DateTimePicker1.Text = "15 Mei 2017" Or DateTimePicker1.Text = "16 Mei 2017" Or DateTimePicker1.Text = "17 Mei 2017" Or DateTimePicker1.Text = "18 Mei 2017" Or DateTimePicker1.Text = "19 Mei 2017" Or DateTimePicker1.Text = "20 Mei 2017" Or DateTimePicker1.Text = "21 Mei 2017" Or DateTimePicker1.Text = "22 Mei 2017" Or DateTimePicker1.Text = "23 Mei 2017" Or DateTimePicker1.Text = "24 Mei 2017" Or DateTimePicker1.Text = "25 Mei 2017" Or DateTimePicker1.Text = "26 Mei 2017" Or DateTimePicker1.Text = "27 Mei 2017" Or DateTimePicker1.Text = "28 Mei 2017" Or DateTimePicker1.Text = "29 Mei 2017" Or DateTimePicker1.Text = "30 Mei 2017" Or DateTimePicker1.Text = "31 Mei 2017" Then
            da = New OleDbDataAdapter("select * from rplb_mei", conn)
            ds = New DataSet
            da.Fill(ds)
            dgv2.DataSource = ds.Tables(0)
        ElseIf DateTimePicker1.Text = "01 Juni 2017" Or DateTimePicker1.Text = "02 Juni 2017" Or DateTimePicker1.Text = "03 Juni 2017" Or DateTimePicker1.Text = "04 Juni 2017" Or DateTimePicker1.Text = "05 Juni 2017" Or DateTimePicker1.Text = "06 Juni 2017" Or DateTimePicker1.Text = "07 Juni 2017" Or DateTimePicker1.Text = "08 Juni 2017" Or DateTimePicker1.Text = "09 Juni 2017" Or DateTimePicker1.Text = "10 Juni 2017" Or DateTimePicker1.Text = "11 Juni 2017" Or DateTimePicker1.Text = "12 Juni 2017" Or DateTimePicker1.Text = "13 Juni 2017" Or DateTimePicker1.Text = "14 Juni 2017" Or DateTimePicker1.Text = "15 Juni 2017" Or DateTimePicker1.Text = "16 Juni 2017" Or DateTimePicker1.Text = "17 Juni 2017" Or DateTimePicker1.Text = "18 Juni 2017" Or DateTimePicker1.Text = "19 Juni 2017" Or DateTimePicker1.Text = "20 Juni 2017" Or DateTimePicker1.Text = "21 Juni 2017" Or DateTimePicker1.Text = "22 Juni 2017" Or DateTimePicker1.Text = "23 Juni 2017" Or DateTimePicker1.Text = "24 Juni 2017" Or DateTimePicker1.Text = "25 Juni 2017" Or DateTimePicker1.Text = "26 Juni 2017" Or DateTimePicker1.Text = "27 Juni 2017" Or DateTimePicker1.Text = "28 Juni 2017" Or DateTimePicker1.Text = "29 Juni 2017" Or DateTimePicker1.Text = "30 Juni 2017" Then
            da = New OleDbDataAdapter("select * from rplb_juni", conn)
            ds = New DataSet
            da.Fill(ds)
            dgv2.DataSource = ds.Tables(0)

        End If
    End Sub

    Private Sub Panel1_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs)

    End Sub

    Private Sub Panel1_Paint_1(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles Panel1.Paint
        Call jumlah()
    End Sub

    Private Sub Panel2_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles Panel2.Paint
        Call tampilgrid1()
    End Sub

    Private Sub Button8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button8.Click
        tambah2.Show()
    End Sub

    Private Sub Button2_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        tambah.Show()
    End Sub

    Private Sub Button3_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        edit.Show()
    End Sub

    Private Sub Button4_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        If edit.TextBox1.Text = "" Then
            MsgBox("Data ada yang kosong")
            Exit Sub
        End If
        Call carikode()
        If Not dr.HasRows Then
            MsgBox("NIS yang anda masukkan tidak terdaftar")
            Exit Sub
        End If
        If MessageBox.Show("Yakin ingin hapus data ini...??", "", MessageBoxButtons.YesNo) = Windows.Forms.DialogResult.Yes Then
            Call koneksi()
            Dim hapus As String = "delete from rplb where NIS= '" & edit.TextBox1.Text & "'"
            cmd = New OleDbCommand(hapus, conn)
            cmd.ExecuteNonQuery()

            Call tampilgrid()
        End If
    End Sub

    Private Sub dgv_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgv.CellContentClick
        On Error Resume Next
        edit.TextBox1.Text = dgv.Rows(e.RowIndex).Cells(0).Value
        Call carikode()
        If dr.HasRows Then
            Call ketemu()
        End If
    End Sub

    Private Sub dgv2_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgv2.CellContentClick

    End Sub

    Private Sub dgv2_CellMouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles dgv2.CellMouseClick
        On Error Resume Next
        TextBox3.Text = dgv2.Rows(e.RowIndex).Cells(0).Value
        edit2.TextBox1.Text = dgv2.Rows(e.RowIndex).Cells(0).Value
        Call carikode2()
        If dr.HasRows Then
            Call edit2.ketemu3()
            Call ketemu2()

            Dim hasil, tgl1, tgl2, tgl3, tgl4, tgl5, tgl6, tgl7, tgl8, tgl9, tgl10, tgl11, tgl12, tgl13, tgl14, tgl15, tgl16, tgl17, tgl18, tgl19, tgl20, tgl21, tgl22, tgl23, tgl24, tgl25, tgl26, tgl27, tgl28, tgl29, tgl30, tgl31 As Integer

            If TextBox3.Text = dgv2.Rows(e.RowIndex).Cells(0).Value Then
                Call carikode2()
                If dr.Item(2) Is DBNull.Value Or dr.Item(2) = "Sakit" Or dr.Item(2) = "Izin" Or dr.Item(2) = "Alfa" Then
                    tgl1 = "0"
                ElseIf dr.Item(2) = "Hadir" Then
                    tgl1 = "1"
                End If
                If dr.Item(3) Is DBNull.Value Or dr.Item(3) = "Sakit" Or dr.Item(3) = "Izin" Or dr.Item(3) = "Alfa" Then
                    tgl2 = "0"
                ElseIf dr.Item(3) = "Hadir" Then
                    tgl2 = "1"
                End If
                If dr.Item(4) Is DBNull.Value Or dr.Item(4) = "Sakit" Or dr.Item(4) = "Izin" Or dr.Item(4) = "Alfa" Then
                    tgl3 = "0"
                ElseIf dr.Item(4) = "Hadir" Then
                    tgl3 = "1"
                End If
                If dr.Item(5) Is DBNull.Value Or dr.Item(5) = "Sakit" Or dr.Item(5) = "Izin" Or dr.Item(5) = "Alfa" Then
                    tgl4 = "0"
                ElseIf dr.Item(5) = "Hadir" Then
                    tgl4 = "1"
                End If
                If dr.Item(6) Is DBNull.Value Or dr.Item(6) = "Sakit" Or dr.Item(6) = "Izin" Or dr.Item(6) = "Alfa" Then
                    tgl5 = "0"
                ElseIf dr.Item(6) = "Hadir" Then
                    tgl5 = "1"
                End If
                If dr.Item(7) Is DBNull.Value Or dr.Item(7) = "Sakit" Or dr.Item(7) = "Izin" Or dr.Item(7) = "Alfa" Then
                    tgl6 = "0"
                ElseIf dr.Item(7) = "Hadir" Then
                    tgl6 = "1"
                End If
                If dr.Item(8) Is DBNull.Value Or dr.Item(8) = "Sakit" Or dr.Item(8) = "Izin" Or dr.Item(8) = "Alfa" Then
                    tgl7 = "0"
                ElseIf dr.Item(8) = "Hadir" Then
                    tgl7 = "1"
                End If
                If dr.Item(9) Is DBNull.Value Or dr.Item(9) = "Sakit" Or dr.Item(9) = "Izin" Or dr.Item(9) = "Alfa" Then
                    tgl8 = "0"
                ElseIf dr.Item(9) = "Hadir" Then
                    tgl8 = "1"
                End If
                If dr.Item(10) Is DBNull.Value Or dr.Item(10) = "Sakit" Or dr.Item(10) = "Izin" Or dr.Item(10) = "Alfa" Then
                    tgl9 = "0"
                ElseIf dr.Item(10) = "Hadir" Then
                    tgl9 = "1"
                End If
                If dr.Item(11) Is DBNull.Value Or dr.Item(11) = "Sakit" Or dr.Item(11) = "Izin" Or dr.Item(11) = "Alfa" Then
                    tgl10 = "0"
                ElseIf dr.Item(11) = "Hadir" Then
                    tgl10 = "1"
                End If
                If dr.Item(12) Is DBNull.Value Or dr.Item(12) = "Sakit" Or dr.Item(12) = "Izin" Or dr.Item(12) = "Alfa" Then
                    tgl11 = "0"
                ElseIf dr.Item(12) = "Hadir" Then
                    tgl11 = "1"
                End If
                If dr.Item(13) Is DBNull.Value Or dr.Item(13) = "Sakit" Or dr.Item(13) = "Izin" Or dr.Item(13) = "Alfa" Then
                    tgl12 = "0"
                ElseIf dr.Item(13) = "Hadir" Then
                    tgl12 = "1"
                End If
                If dr.Item(14) Is DBNull.Value Or dr.Item(14) = "Sakit" Or dr.Item(14) = "Izin" Or dr.Item(14) = "Alfa" Then
                    tgl13 = "0"
                ElseIf dr.Item(14) = "Hadir" Then
                    tgl13 = "1"
                End If
                If dr.Item(15) Is DBNull.Value Or dr.Item(15) = "Sakit" Or dr.Item(15) = "Izin" Or dr.Item(15) = "Alfa" Then
                    tgl14 = "0"
                ElseIf dr.Item(15) = "Hadir" Then
                    tgl14 = "1"
                End If
                If dr.Item(16) Is DBNull.Value Or dr.Item(16) = "Sakit" Or dr.Item(16) = "Izin" Or dr.Item(16) = "Alfa" Then
                    tgl15 = "0"
                ElseIf dr.Item(16) = "Hadir" Then
                    tgl15 = "1"
                End If
                If dr.Item(17) Is DBNull.Value Or dr.Item(17) = "Sakit" Or dr.Item(17) = "Izin" Or dr.Item(17) = "Alfa" Then
                    tgl16 = "0"
                ElseIf dr.Item(17) = "Hadir" Then
                    tgl16 = "1"
                End If
                If dr.Item(18) Is DBNull.Value Or dr.Item(18) = "Sakit" Or dr.Item(18) = "Izin" Or dr.Item(18) = "Alfa" Then
                    tgl17 = "0"
                ElseIf dr.Item(18) = "Hadir" Then
                    tgl17 = "1"
                End If
                If dr.Item(19) Is DBNull.Value Or dr.Item(19) = "Sakit" Or dr.Item(19) = "Izin" Or dr.Item(19) = "Alfa" Then
                    tgl18 = "0"
                ElseIf dr.Item(19) = "Hadir" Then
                    tgl18 = "1"
                End If
                If dr.Item(20) Is DBNull.Value Or dr.Item(20) = "Sakit" Or dr.Item(20) = "Izin" Or dr.Item(20) = "Alfa" Then
                    tgl19 = "0"
                ElseIf dr.Item(20) = "Hadir" Then
                    tgl19 = "1"
                End If
                If dr.Item(21) Is DBNull.Value Or dr.Item(21) = "Sakit" Or dr.Item(21) = "Izin" Or dr.Item(21) = "Alfa" Then
                    tgl20 = "0"
                ElseIf dr.Item(21) = "Hadir" Then
                    tgl20 = "1"
                End If
                If dr.Item(22) Is DBNull.Value Or dr.Item(22) = "Sakit" Or dr.Item(22) = "Izin" Or dr.Item(22) = "Alfa" Then
                    tgl21 = "0"
                ElseIf dr.Item(22) = "Hadir" Then
                    tgl21 = "1"
                End If
                If dr.Item(23) Is DBNull.Value Or dr.Item(23) = "Sakit" Or dr.Item(23) = "Izin" Or dr.Item(23) = "Alfa" Then
                    tgl22 = "0"
                ElseIf dr.Item(23) = "Hadir" Then
                    tgl22 = "1"
                End If
                If dr.Item(24) Is DBNull.Value Or dr.Item(24) = "Sakit" Or dr.Item(24) = "Izin" Or dr.Item(24) = "Alfa" Then
                    tgl23 = "0"
                ElseIf dr.Item(24) = "Hadir" Then
                    tgl23 = "1"
                End If
                If dr.Item(25) Is DBNull.Value Or dr.Item(25) = "Sakit" Or dr.Item(25) = "Izin" Or dr.Item(25) = "Alfa" Then
                    tgl24 = "0"
                ElseIf dr.Item(25) = "Hadir" Then
                    tgl24 = "1"
                End If
                If dr.Item(26) Is DBNull.Value Or dr.Item(26) = "Sakit" Or dr.Item(26) = "Izin" Or dr.Item(26) = "Alfa" Then
                    tgl25 = "0"
                ElseIf dr.Item(26) = "Hadir" Then
                    tgl25 = "1"
                End If
                If dr.Item(27) Is DBNull.Value Or dr.Item(27) = "Sakit" Or dr.Item(27) = "Izin" Or dr.Item(27) = "Alfa" Then
                    tgl26 = "0"
                ElseIf dr.Item(27) = "Hadir" Then
                    tgl26 = "1"
                End If
                If dr.Item(28) Is DBNull.Value Or dr.Item(28) = "Sakit" Or dr.Item(28) = "Izin" Or dr.Item(28) = "Alfa" Then
                    tgl27 = "0"
                ElseIf dr.Item(28) = "Hadir" Then
                    tgl27 = "1"
                End If
                If dr.Item(29) Is DBNull.Value Or dr.Item(29) = "Sakit" Or dr.Item(29) = "Izin" Or dr.Item(29) = "Alfa" Then
                    tgl28 = "0"
                ElseIf dr.Item(29) = "Hadir" Then
                    tgl28 = "1"
                End If
                If dr.Item(30) Is DBNull.Value Or dr.Item(30) = "Sakit" Or dr.Item(30) = "Izin" Or dr.Item(30) = "Alfa" Then
                    tgl29 = "0"
                ElseIf dr.Item(30) = "Hadir" Then
                    tgl29 = "1"
                End If
                If dr.Item(31) Is DBNull.Value Or dr.Item(31) = "Sakit" Or dr.Item(31) = "Izin" Or dr.Item(31) = "Alfa" Then
                    tgl30 = "0"
                ElseIf dr.Item(31) = "Hadir" Then
                    tgl30 = "1"
                End If
                If dr.Item(32) Is DBNull.Value Or dr.Item(32) = "Sakit" Or dr.Item(32) = "Izin" Or dr.Item(32) = "Alfa" Then
                    tgl31 = "0"
                ElseIf dr.Item(32) = "Hadir" Then
                    tgl31 = "1"
                End If
                hasil = tgl1 + tgl2 + tgl3 + tgl4 + tgl5 + tgl6 + tgl7 + tgl8 + tgl9 + tgl10 + tgl11 + tgl12 + tgl13 + tgl14 + tgl15 + tgl16 + tgl17 + tgl18 + tgl19 + tgl20 + tgl21 + tgl22 + tgl23 + tgl24 + tgl25 + tgl26 + tgl27 + tgl28 + tgl29 + tgl30 + tgl31
                Label15.Text = hasil
            End If

            If TextBox3.Text = dgv2.Rows(e.RowIndex).Cells(0).Value Then
                Call carikode2()
                If dr.Item(2) Is DBNull.Value Or dr.Item(2) = "Hadir" Or dr.Item(2) = "Izin" Or dr.Item(2) = "Alfa" Then
                    tgl1 = "0"
                ElseIf dr.Item(2) = "Sakit" Then
                    tgl1 = "1"
                End If
                If dr.Item(3) Is DBNull.Value Or dr.Item(3) = "Hadir" Or dr.Item(3) = "Izin" Or dr.Item(3) = "Alfa" Then
                    tgl2 = "0"
                ElseIf dr.Item(3) = "Sakit" Then
                    tgl2 = "1"
                End If
                If dr.Item(4) Is DBNull.Value Or dr.Item(4) = "Hadir" Or dr.Item(4) = "Izin" Or dr.Item(4) = "Alfa" Then
                    tgl3 = "0"
                ElseIf dr.Item(4) = "Sakit" Then
                    tgl3 = "1"
                End If
                If dr.Item(5) Is DBNull.Value Or dr.Item(5) = "Hadir" Or dr.Item(5) = "Izin" Or dr.Item(5) = "Alfa" Then
                    tgl4 = "0"
                ElseIf dr.Item(5) = "Sakit" Then
                    tgl4 = "1"
                End If
                If dr.Item(6) Is DBNull.Value Or dr.Item(6) = "Hadir" Or dr.Item(6) = "Izin" Or dr.Item(6) = "Alfa" Then
                    tgl5 = "0"
                ElseIf dr.Item(6) = "Sakit" Then
                    tgl5 = "1"
                End If
                If dr.Item(7) Is DBNull.Value Or dr.Item(7) = "Hadir" Or dr.Item(7) = "Izin" Or dr.Item(7) = "Alfa" Then
                    tgl6 = "0"
                ElseIf dr.Item(7) = "Sakit" Then
                    tgl6 = "1"
                End If
                If dr.Item(8) Is DBNull.Value Or dr.Item(8) = "Hadir" Or dr.Item(8) = "Izin" Or dr.Item(8) = "Alfa" Then
                    tgl7 = "0"
                ElseIf dr.Item(8) = "Sakit" Then
                    tgl7 = "1"
                End If
                If dr.Item(9) Is DBNull.Value Or dr.Item(9) = "Hadir" Or dr.Item(9) = "Izin" Or dr.Item(9) = "Alfa" Then
                    tgl8 = "0"
                ElseIf dr.Item(9) = "Sakit" Then
                    tgl8 = "1"
                End If
                If dr.Item(10) Is DBNull.Value Or dr.Item(10) = "Hadir" Or dr.Item(10) = "Izin" Or dr.Item(10) = "Alfa" Then
                    tgl9 = "0"
                ElseIf dr.Item(10) = "Sakit" Then
                    tgl9 = "1"
                End If
                If dr.Item(11) Is DBNull.Value Or dr.Item(11) = "Hadir" Or dr.Item(11) = "Izin" Or dr.Item(11) = "Alfa" Then
                    tgl10 = "0"
                ElseIf dr.Item(11) = "Sakit" Then
                    tgl10 = "1"
                End If
                If dr.Item(12) Is DBNull.Value Or dr.Item(12) = "Hadir" Or dr.Item(12) = "Izin" Or dr.Item(12) = "Alfa" Then
                    tgl11 = "0"
                ElseIf dr.Item(12) = "Sakit" Then
                    tgl11 = "1"
                End If
                If dr.Item(13) Is DBNull.Value Or dr.Item(13) = "Hadir" Or dr.Item(13) = "Izin" Or dr.Item(13) = "Alfa" Then
                    tgl12 = "0"
                ElseIf dr.Item(13) = "Sakit" Then
                    tgl12 = "1"
                End If
                If dr.Item(14) Is DBNull.Value Or dr.Item(14) = "Hadir" Or dr.Item(14) = "Izin" Or dr.Item(14) = "Alfa" Then
                    tgl13 = "0"
                ElseIf dr.Item(14) = "Sakit" Then
                    tgl13 = "1"
                End If
                If dr.Item(15) Is DBNull.Value Or dr.Item(15) = "Hadir" Or dr.Item(15) = "Izin" Or dr.Item(15) = "Alfa" Then
                    tgl14 = "0"
                ElseIf dr.Item(15) = "Sakit" Then
                    tgl14 = "1"
                End If
                If dr.Item(16) Is DBNull.Value Or dr.Item(16) = "Hadir" Or dr.Item(16) = "Izin" Or dr.Item(16) = "Alfa" Then
                    tgl15 = "0"
                ElseIf dr.Item(16) = "Sakit" Then
                    tgl15 = "1"
                End If
                If dr.Item(17) Is DBNull.Value Or dr.Item(17) = "Hadir" Or dr.Item(17) = "Izin" Or dr.Item(17) = "Alfa" Then
                    tgl16 = "0"
                ElseIf dr.Item(17) = "Sakit" Then
                    tgl16 = "1"
                End If
                If dr.Item(18) Is DBNull.Value Or dr.Item(18) = "Hadir" Or dr.Item(18) = "Izin" Or dr.Item(18) = "Alfa" Then
                    tgl17 = "0"
                ElseIf dr.Item(18) = "Sakit" Then
                    tgl17 = "1"
                End If
                If dr.Item(19) Is DBNull.Value Or dr.Item(19) = "Hadir" Or dr.Item(19) = "Izin" Or dr.Item(19) = "Alfa" Then
                    tgl18 = "0"
                ElseIf dr.Item(19) = "Sakit" Then
                    tgl18 = "1"
                End If
                If dr.Item(20) Is DBNull.Value Or dr.Item(20) = "Hadir" Or dr.Item(20) = "Izin" Or dr.Item(20) = "Alfa" Then
                    tgl19 = "0"
                ElseIf dr.Item(20) = "Sakit" Then
                    tgl19 = "1"
                End If
                If dr.Item(21) Is DBNull.Value Or dr.Item(21) = "Hadir" Or dr.Item(21) = "Izin" Or dr.Item(21) = "Alfa" Then
                    tgl20 = "0"
                ElseIf dr.Item(21) = "Sakit" Then
                    tgl20 = "1"
                End If
                If dr.Item(22) Is DBNull.Value Or dr.Item(22) = "Hadir" Or dr.Item(22) = "Izin" Or dr.Item(22) = "Alfa" Then
                    tgl21 = "0"
                ElseIf dr.Item(22) = "Sakit" Then
                    tgl21 = "1"
                End If
                If dr.Item(23) Is DBNull.Value Or dr.Item(23) = "Hadir" Or dr.Item(23) = "Izin" Or dr.Item(23) = "Alfa" Then
                    tgl22 = "0"
                ElseIf dr.Item(23) = "Sakit" Then
                    tgl22 = "1"
                End If
                If dr.Item(24) Is DBNull.Value Or dr.Item(24) = "Hadir" Or dr.Item(24) = "Izin" Or dr.Item(24) = "Alfa" Then
                    tgl23 = "0"
                ElseIf dr.Item(24) = "Sakit" Then
                    tgl23 = "1"
                End If
                If dr.Item(25) Is DBNull.Value Or dr.Item(25) = "Hadir" Or dr.Item(25) = "Izin" Or dr.Item(25) = "Alfa" Then
                    tgl24 = "0"
                ElseIf dr.Item(25) = "Sakit" Then
                    tgl24 = "1"
                End If
                If dr.Item(26) Is DBNull.Value Or dr.Item(26) = "Hadir" Or dr.Item(26) = "Izin" Or dr.Item(26) = "Alfa" Then
                    tgl25 = "0"
                ElseIf dr.Item(26) = "Sakit" Then
                    tgl25 = "1"
                End If
                If dr.Item(27) Is DBNull.Value Or dr.Item(27) = "Hadir" Or dr.Item(27) = "Izin" Or dr.Item(27) = "Alfa" Then
                    tgl26 = "0"
                ElseIf dr.Item(27) = "Sakit" Then
                    tgl26 = "1"
                End If
                If dr.Item(28) Is DBNull.Value Or dr.Item(28) = "Hadir" Or dr.Item(28) = "Izin" Or dr.Item(28) = "Alfa" Then
                    tgl27 = "0"
                ElseIf dr.Item(28) = "Sakit" Then
                    tgl27 = "1"
                End If
                If dr.Item(29) Is DBNull.Value Or dr.Item(29) = "Hadir" Or dr.Item(29) = "Izin" Or dr.Item(29) = "Alfa" Then
                    tgl28 = "0"
                ElseIf dr.Item(29) = "Sakit" Then
                    tgl28 = "1"
                End If
                If dr.Item(30) Is DBNull.Value Or dr.Item(30) = "Hadir" Or dr.Item(30) = "Izin" Or dr.Item(30) = "Alfa" Then
                    tgl29 = "0"
                ElseIf dr.Item(30) = "Sakit" Then
                    tgl29 = "1"
                End If
                If dr.Item(31) Is DBNull.Value Or dr.Item(31) = "Hadir" Or dr.Item(31) = "Izin" Or dr.Item(31) = "Alfa" Then
                    tgl30 = "0"
                ElseIf dr.Item(31) = "Sakit" Then
                    tgl30 = "1"
                End If
                If dr.Item(32) Is DBNull.Value Or dr.Item(32) = "Hadir" Or dr.Item(32) = "Izin" Or dr.Item(32) = "Alfa" Then
                    tgl31 = "0"
                ElseIf dr.Item(32) = "Sakit" Then
                    tgl31 = "1"
                End If
                hasil = tgl1 + tgl2 + tgl3 + tgl4 + tgl5 + tgl6 + tgl7 + tgl8 + tgl9 + tgl10 + tgl11 + tgl12 + tgl13 + tgl14 + tgl15 + tgl16 + tgl17 + tgl18 + tgl19 + tgl20 + tgl21 + tgl22 + tgl23 + tgl24 + tgl25 + tgl26 + tgl27 + tgl28 + tgl29 + tgl30 + tgl31
                Label9.Text = hasil
            End If

            If TextBox3.Text = dgv2.Rows(e.RowIndex).Cells(0).Value Then
                Call carikode2()
                If dr.Item(2) Is DBNull.Value Or dr.Item(2) = "Hadir" Or dr.Item(2) = "Sakit" Or dr.Item(2) = "Alfa" Then
                    tgl1 = "0"
                ElseIf dr.Item(2) = "Izin" Then
                    tgl1 = "1"
                End If
                If dr.Item(3) Is DBNull.Value Or dr.Item(3) = "Hadir" Or dr.Item(3) = "Sakit" Or dr.Item(3) = "Alfa" Then
                    tgl2 = "0"
                ElseIf dr.Item(3) = "Izin" Then
                    tgl2 = "1"
                End If
                If dr.Item(4) Is DBNull.Value Or dr.Item(4) = "Hadir" Or dr.Item(4) = "Sakit" Or dr.Item(4) = "Alfa" Then
                    tgl3 = "0"
                ElseIf dr.Item(4) = "Izin" Then
                    tgl3 = "1"
                End If
                If dr.Item(5) Is DBNull.Value Or dr.Item(5) = "Hadir" Or dr.Item(5) = "Sakit" Or dr.Item(5) = "Alfa" Then
                    tgl4 = "0"
                ElseIf dr.Item(5) = "Izin" Then
                    tgl4 = "1"
                End If
                If dr.Item(6) Is DBNull.Value Or dr.Item(6) = "Hadir" Or dr.Item(6) = "Sakit" Or dr.Item(6) = "Alfa" Then
                    tgl5 = "0"
                ElseIf dr.Item(6) = "Izin" Then
                    tgl5 = "1"
                End If
                If dr.Item(7) Is DBNull.Value Or dr.Item(7) = "Hadir" Or dr.Item(7) = "Sakit" Or dr.Item(7) = "Alfa" Then
                    tgl6 = "0"
                ElseIf dr.Item(7) = "Izin" Then
                    tgl6 = "1"
                End If
                If dr.Item(8) Is DBNull.Value Or dr.Item(8) = "Hadir" Or dr.Item(8) = "Sakit" Or dr.Item(8) = "Alfa" Then
                    tgl7 = "0"
                ElseIf dr.Item(8) = "Izin" Then
                    tgl7 = "1"
                End If
                If dr.Item(9) Is DBNull.Value Or dr.Item(9) = "Hadir" Or dr.Item(9) = "Sakit" Or dr.Item(9) = "Alfa" Then
                    tgl8 = "0"
                ElseIf dr.Item(9) = "Izin" Then
                    tgl8 = "1"
                End If
                If dr.Item(10) Is DBNull.Value Or dr.Item(10) = "Hadir" Or dr.Item(10) = "Sakit" Or dr.Item(10) = "Alfa" Then
                    tgl9 = "0"
                ElseIf dr.Item(10) = "Izin" Then
                    tgl9 = "1"
                End If
                If dr.Item(11) Is DBNull.Value Or dr.Item(11) = "Hadir" Or dr.Item(11) = "Sakit" Or dr.Item(11) = "Alfa" Then
                    tgl10 = "0"
                ElseIf dr.Item(11) = "Izin" Then
                    tgl10 = "1"
                End If
                If dr.Item(12) Is DBNull.Value Or dr.Item(12) = "Hadir" Or dr.Item(12) = "Sakit" Or dr.Item(12) = "Alfa" Then
                    tgl11 = "0"
                ElseIf dr.Item(12) = "Izin" Then
                    tgl11 = "1"
                End If
                If dr.Item(13) Is DBNull.Value Or dr.Item(13) = "Hadir" Or dr.Item(13) = "Sakit" Or dr.Item(13) = "Alfa" Then
                    tgl12 = "0"
                ElseIf dr.Item(13) = "Izin" Then
                    tgl12 = "1"
                End If
                If dr.Item(14) Is DBNull.Value Or dr.Item(14) = "Hadir" Or dr.Item(14) = "Sakit" Or dr.Item(14) = "Alfa" Then
                    tgl13 = "0"
                ElseIf dr.Item(14) = "Izin" Then
                    tgl13 = "1"
                End If
                If dr.Item(15) Is DBNull.Value Or dr.Item(15) = "Hadir" Or dr.Item(15) = "Sakit" Or dr.Item(15) = "Alfa" Then
                    tgl14 = "0"
                ElseIf dr.Item(15) = "Izin" Then
                    tgl14 = "1"
                End If
                If dr.Item(16) Is DBNull.Value Or dr.Item(16) = "Hadir" Or dr.Item(16) = "Sakit" Or dr.Item(16) = "Alfa" Then
                    tgl15 = "0"
                ElseIf dr.Item(16) = "Izin" Then
                    tgl15 = "1"
                End If
                If dr.Item(17) Is DBNull.Value Or dr.Item(17) = "Hadir" Or dr.Item(17) = "Sakit" Or dr.Item(17) = "Alfa" Then
                    tgl16 = "0"
                ElseIf dr.Item(17) = "Izin" Then
                    tgl16 = "1"
                End If
                If dr.Item(18) Is DBNull.Value Or dr.Item(18) = "Hadir" Or dr.Item(18) = "Sakit" Or dr.Item(18) = "Alfa" Then
                    tgl17 = "0"
                ElseIf dr.Item(18) = "Izin" Then
                    tgl17 = "1"
                End If
                If dr.Item(19) Is DBNull.Value Or dr.Item(19) = "Hadir" Or dr.Item(19) = "Sakit" Or dr.Item(19) = "Alfa" Then
                    tgl18 = "0"
                ElseIf dr.Item(19) = "Izin" Then
                    tgl18 = "1"
                End If
                If dr.Item(20) Is DBNull.Value Or dr.Item(20) = "Hadir" Or dr.Item(20) = "Sakit" Or dr.Item(20) = "Alfa" Then
                    tgl19 = "0"
                ElseIf dr.Item(20) = "Izin" Then
                    tgl19 = "1"
                End If
                If dr.Item(21) Is DBNull.Value Or dr.Item(21) = "Hadir" Or dr.Item(21) = "Sakit" Or dr.Item(21) = "Alfa" Then
                    tgl20 = "0"
                ElseIf dr.Item(21) = "Izin" Then
                    tgl20 = "1"
                End If
                If dr.Item(22) Is DBNull.Value Or dr.Item(22) = "Hadir" Or dr.Item(22) = "Sakit" Or dr.Item(22) = "Alfa" Then
                    tgl21 = "0"
                ElseIf dr.Item(22) = "Izin" Then
                    tgl21 = "1"
                End If
                If dr.Item(23) Is DBNull.Value Or dr.Item(23) = "Hadir" Or dr.Item(23) = "Sakit" Or dr.Item(23) = "Alfa" Then
                    tgl22 = "0"
                ElseIf dr.Item(23) = "Izin" Then
                    tgl22 = "1"
                End If
                If dr.Item(24) Is DBNull.Value Or dr.Item(24) = "Hadir" Or dr.Item(24) = "Sakit" Or dr.Item(24) = "Alfa" Then
                    tgl23 = "0"
                ElseIf dr.Item(24) = "Izin" Then
                    tgl23 = "1"
                End If
                If dr.Item(25) Is DBNull.Value Or dr.Item(25) = "Hadir" Or dr.Item(25) = "Sakit" Or dr.Item(25) = "Alfa" Then
                    tgl24 = "0"
                ElseIf dr.Item(25) = "Izin" Then
                    tgl24 = "1"
                End If
                If dr.Item(26) Is DBNull.Value Or dr.Item(26) = "Hadir" Or dr.Item(26) = "Sakit" Or dr.Item(26) = "Alfa" Then
                    tgl25 = "0"
                ElseIf dr.Item(26) = "Izin" Then
                    tgl25 = "1"
                End If
                If dr.Item(27) Is DBNull.Value Or dr.Item(27) = "Hadir" Or dr.Item(27) = "Sakit" Or dr.Item(27) = "Alfa" Then
                    tgl26 = "0"
                ElseIf dr.Item(27) = "Izin" Then
                    tgl26 = "1"
                End If
                If dr.Item(28) Is DBNull.Value Or dr.Item(28) = "Hadir" Or dr.Item(28) = "Sakit" Or dr.Item(28) = "Alfa" Then
                    tgl27 = "0"
                ElseIf dr.Item(28) = "Izin" Then
                    tgl27 = "1"
                End If
                If dr.Item(29) Is DBNull.Value Or dr.Item(29) = "Hadir" Or dr.Item(29) = "Sakit" Or dr.Item(29) = "Alfa" Then
                    tgl28 = "0"
                ElseIf dr.Item(29) = "Izin" Then
                    tgl28 = "1"
                End If
                If dr.Item(30) Is DBNull.Value Or dr.Item(30) = "Hadir" Or dr.Item(30) = "Sakit" Or dr.Item(30) = "Alfa" Then
                    tgl29 = "0"
                ElseIf dr.Item(30) = "Izin" Then
                    tgl29 = "1"
                End If
                If dr.Item(31) Is DBNull.Value Or dr.Item(31) = "Hadir" Or dr.Item(31) = "Sakit" Or dr.Item(31) = "Alfa" Then
                    tgl30 = "0"
                ElseIf dr.Item(31) = "Izin" Then
                    tgl30 = "1"
                End If
                If dr.Item(32) Is DBNull.Value Or dr.Item(32) = "Hadir" Or dr.Item(32) = "Sakit" Or dr.Item(32) = "Alfa" Then
                    tgl31 = "0"
                ElseIf dr.Item(32) = "Izin" Then
                    tgl31 = "1"
                End If
                hasil = tgl1 + tgl2 + tgl3 + tgl4 + tgl5 + tgl6 + tgl7 + tgl8 + tgl9 + tgl10 + tgl11 + tgl12 + tgl13 + tgl14 + tgl15 + tgl16 + tgl17 + tgl18 + tgl19 + tgl20 + tgl21 + tgl22 + tgl23 + tgl24 + tgl25 + tgl26 + tgl27 + tgl28 + tgl29 + tgl30 + tgl31
                Label11.Text = hasil
            End If


            If TextBox3.Text = dgv2.Rows(e.RowIndex).Cells(0).Value Then
                Call carikode2()
                If dr.Item(2) Is DBNull.Value Or dr.Item(2) = "Hadir" Or dr.Item(2) = "Sakit" Or dr.Item(2) = "Izin" Then
                    tgl1 = "0"
                ElseIf dr.Item(2) = "Alfa" Then
                    tgl1 = "1"
                End If
                If dr.Item(3) Is DBNull.Value Or dr.Item(3) = "Hadir" Or dr.Item(3) = "Sakit" Or dr.Item(3) = "Izin" Then
                    tgl2 = "0"
                ElseIf dr.Item(3) = "Alfa" Then
                    tgl2 = "1"
                End If
                If dr.Item(4) Is DBNull.Value Or dr.Item(4) = "Hadir" Or dr.Item(4) = "Sakit" Or dr.Item(4) = "Izin" Then
                    tgl3 = "0"
                ElseIf dr.Item(4) = "Alfa" Then
                    tgl3 = "1"
                End If
                If dr.Item(5) Is DBNull.Value Or dr.Item(5) = "Hadir" Or dr.Item(5) = "Sakit" Or dr.Item(5) = "Izin" Then
                    tgl4 = "0"
                ElseIf dr.Item(5) = "Alfa" Then
                    tgl4 = "1"
                End If
                If dr.Item(6) Is DBNull.Value Or dr.Item(6) = "Hadir" Or dr.Item(6) = "Sakit" Or dr.Item(6) = "Izin" Then
                    tgl5 = "0"
                ElseIf dr.Item(6) = "Alfa" Then
                    tgl5 = "1"
                End If
                If dr.Item(7) Is DBNull.Value Or dr.Item(7) = "Hadir" Or dr.Item(7) = "Sakit" Or dr.Item(7) = "Izin" Then
                    tgl6 = "0"
                ElseIf dr.Item(7) = "Alfa" Then
                    tgl6 = "1"
                End If
                If dr.Item(8) Is DBNull.Value Or dr.Item(8) = "Hadir" Or dr.Item(8) = "Sakit" Or dr.Item(8) = "Izin" Then
                    tgl7 = "0"
                ElseIf dr.Item(8) = "Alfa" Then
                    tgl7 = "1"
                End If
                If dr.Item(9) Is DBNull.Value Or dr.Item(9) = "Hadir" Or dr.Item(9) = "Sakit" Or dr.Item(9) = "Izin" Then
                    tgl8 = "0"
                ElseIf dr.Item(9) = "Alfa" Then
                    tgl8 = "1"
                End If
                If dr.Item(10) Is DBNull.Value Or dr.Item(10) = "Hadir" Or dr.Item(10) = "Sakit" Or dr.Item(10) = "Izin" Then
                    tgl9 = "0"
                ElseIf dr.Item(10) = "Alfa" Then
                    tgl9 = "1"
                End If
                If dr.Item(11) Is DBNull.Value Or dr.Item(11) = "Hadir" Or dr.Item(11) = "Sakit" Or dr.Item(11) = "Izin" Then
                    tgl10 = "0"
                ElseIf dr.Item(11) = "Alfa" Then
                    tgl10 = "1"
                End If
                If dr.Item(12) Is DBNull.Value Or dr.Item(12) = "Hadir" Or dr.Item(12) = "Sakit" Or dr.Item(12) = "Izin" Then
                    tgl11 = "0"
                ElseIf dr.Item(12) = "Alfa" Then
                    tgl11 = "1"
                End If
                If dr.Item(13) Is DBNull.Value Or dr.Item(13) = "Hadir" Or dr.Item(13) = "Sakit" Or dr.Item(13) = "Izin" Then
                    tgl12 = "0"
                ElseIf dr.Item(13) = "Alfa" Then
                    tgl12 = "1"
                End If
                If dr.Item(14) Is DBNull.Value Or dr.Item(14) = "Hadir" Or dr.Item(14) = "Sakit" Or dr.Item(14) = "Izin" Then
                    tgl13 = "0"
                ElseIf dr.Item(14) = "Alfa" Then
                    tgl13 = "1"
                End If
                If dr.Item(15) Is DBNull.Value Or dr.Item(15) = "Hadir" Or dr.Item(15) = "Sakit" Or dr.Item(15) = "Izin" Then
                    tgl14 = "0"
                ElseIf dr.Item(15) = "Alfa" Then
                    tgl14 = "1"
                End If
                If dr.Item(16) Is DBNull.Value Or dr.Item(16) = "Hadir" Or dr.Item(16) = "Sakit" Or dr.Item(16) = "Izin" Then
                    tgl15 = "0"
                ElseIf dr.Item(16) = "Alfa" Then
                    tgl15 = "1"
                End If
                If dr.Item(17) Is DBNull.Value Or dr.Item(17) = "Hadir" Or dr.Item(17) = "Sakit" Or dr.Item(17) = "Izin" Then
                    tgl16 = "0"
                ElseIf dr.Item(17) = "Alfa" Then
                    tgl16 = "1"
                End If
                If dr.Item(18) Is DBNull.Value Or dr.Item(18) = "Hadir" Or dr.Item(18) = "Sakit" Or dr.Item(18) = "Izin" Then
                    tgl17 = "0"
                ElseIf dr.Item(18) = "Alfa" Then
                    tgl17 = "1"
                End If
                If dr.Item(19) Is DBNull.Value Or dr.Item(19) = "Hadir" Or dr.Item(19) = "Sakit" Or dr.Item(19) = "Izin" Then
                    tgl18 = "0"
                ElseIf dr.Item(19) = "Alfa" Then
                    tgl18 = "1"
                End If
                If dr.Item(20) Is DBNull.Value Or dr.Item(20) = "Hadir" Or dr.Item(20) = "Sakit" Or dr.Item(20) = "Izin" Then
                    tgl19 = "0"
                ElseIf dr.Item(20) = "Alfa" Then
                    tgl19 = "1"
                End If
                If dr.Item(21) Is DBNull.Value Or dr.Item(21) = "Hadir" Or dr.Item(21) = "Sakit" Or dr.Item(21) = "Izin" Then
                    tgl20 = "0"
                ElseIf dr.Item(21) = "Alfa" Then
                    tgl20 = "1"
                End If
                If dr.Item(22) Is DBNull.Value Or dr.Item(22) = "Hadir" Or dr.Item(22) = "Sakit" Or dr.Item(22) = "Izin" Then
                    tgl21 = "0"
                ElseIf dr.Item(22) = "Alfa" Then
                    tgl21 = "1"
                End If
                If dr.Item(23) Is DBNull.Value Or dr.Item(23) = "Hadir" Or dr.Item(23) = "Sakit" Or dr.Item(23) = "Izin" Then
                    tgl22 = "0"
                ElseIf dr.Item(23) = "Alfa" Then
                    tgl22 = "1"
                End If
                If dr.Item(24) Is DBNull.Value Or dr.Item(24) = "Hadir" Or dr.Item(24) = "Sakit" Or dr.Item(24) = "Izin" Then
                    tgl23 = "0"
                ElseIf dr.Item(24) = "Alfa" Then
                    tgl23 = "1"
                End If
                If dr.Item(25) Is DBNull.Value Or dr.Item(25) = "Hadir" Or dr.Item(25) = "Sakit" Or dr.Item(25) = "Izin" Then
                    tgl24 = "0"
                ElseIf dr.Item(25) = "Alfa" Then
                    tgl24 = "1"
                End If
                If dr.Item(26) Is DBNull.Value Or dr.Item(26) = "Hadir" Or dr.Item(26) = "Sakit" Or dr.Item(26) = "Izin" Then
                    tgl25 = "0"
                ElseIf dr.Item(26) = "Alfa" Then
                    tgl25 = "1"
                End If
                If dr.Item(27) Is DBNull.Value Or dr.Item(27) = "Hadir" Or dr.Item(27) = "Sakit" Or dr.Item(27) = "Izin" Then
                    tgl26 = "0"
                ElseIf dr.Item(27) = "Alfa" Then
                    tgl26 = "1"
                End If
                If dr.Item(28) Is DBNull.Value Or dr.Item(28) = "Hadir" Or dr.Item(28) = "Sakit" Or dr.Item(28) = "Izin" Then
                    tgl27 = "0"
                ElseIf dr.Item(28) = "Alfa" Then
                    tgl27 = "1"
                End If
                If dr.Item(29) Is DBNull.Value Or dr.Item(29) = "Hadir" Or dr.Item(29) = "Sakit" Or dr.Item(29) = "Izin" Then
                    tgl28 = "0"
                ElseIf dr.Item(29) = "Alfa" Then
                    tgl28 = "1"
                End If
                If dr.Item(30) Is DBNull.Value Or dr.Item(30) = "Hadir" Or dr.Item(30) = "Sakit" Or dr.Item(30) = "Izin" Then
                    tgl29 = "0"
                ElseIf dr.Item(30) = "Alfa" Then
                    tgl29 = "1"
                End If
                If dr.Item(31) Is DBNull.Value Or dr.Item(31) = "Hadir" Or dr.Item(31) = "Sakit" Or dr.Item(31) = "Izin" Then
                    tgl30 = "0"
                ElseIf dr.Item(31) = "Alfa" Then
                    tgl30 = "1"
                End If
                If dr.Item(32) Is DBNull.Value Or dr.Item(32) = "Hadir" Or dr.Item(32) = "Sakit" Or dr.Item(32) = "Izin" Then
                    tgl31 = "0"
                ElseIf dr.Item(32) = "Alfa" Then
                    tgl31 = "1"
                End If
                hasil = tgl1 + tgl2 + tgl3 + tgl4 + tgl5 + tgl6 + tgl7 + tgl8 + tgl9 + tgl10 + tgl11 + tgl12 + tgl13 + tgl14 + tgl15 + tgl16 + tgl17 + tgl18 + tgl19 + tgl20 + tgl21 + tgl22 + tgl23 + tgl24 + tgl25 + tgl26 + tgl27 + tgl28 + tgl29 + tgl30 + tgl31
                Label13.Text = hasil
            End If

        End If
    End Sub

    Private Sub TextBox3_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox3.KeyPress
        Dim KeyAscii As Short = Asc(e.KeyChar)
        On Error Resume Next
        If e.KeyChar = Chr(13) Then
            Dim hasil, tgl1, tgl2, tgl3, tgl4, tgl5, tgl6, tgl7, tgl8, tgl9, tgl10, tgl11, tgl12, tgl13, tgl14, tgl15, tgl16, tgl17, tgl18, tgl19, tgl20, tgl21, tgl22, tgl23, tgl24, tgl25, tgl26, tgl27, tgl28, tgl29, tgl30, tgl31 As Integer

            Call koneksi()
            Call carikode2()
            If dr.HasRows Then
                Call ketemu2()

                If e.KeyChar = Chr(13) Then
                    Call carikode2()
                    If dr.Item(2) Is DBNull.Value Or dr.Item(2) = "Sakit" Or dr.Item(2) = "Izin" Or dr.Item(2) = "Alfa" Then
                        tgl1 = "0"
                    ElseIf dr.Item(2) = "Hadir" Then
                        tgl1 = "1"
                    End If
                    If dr.Item(3) Is DBNull.Value Or dr.Item(3) = "Sakit" Or dr.Item(3) = "Izin" Or dr.Item(3) = "Alfa" Then
                        tgl2 = "0"
                    ElseIf dr.Item(3) = "Hadir" Then
                        tgl2 = "1"
                    End If
                    If dr.Item(4) Is DBNull.Value Or dr.Item(4) = "Sakit" Or dr.Item(4) = "Izin" Or dr.Item(4) = "Alfa" Then
                        tgl3 = "0"
                    ElseIf dr.Item(4) = "Hadir" Then
                        tgl3 = "1"
                    End If
                    If dr.Item(5) Is DBNull.Value Or dr.Item(5) = "Sakit" Or dr.Item(5) = "Izin" Or dr.Item(5) = "Alfa" Then
                        tgl4 = "0"
                    ElseIf dr.Item(5) = "Hadir" Then
                        tgl4 = "1"
                    End If
                    If dr.Item(6) Is DBNull.Value Or dr.Item(6) = "Sakit" Or dr.Item(6) = "Izin" Or dr.Item(6) = "Alfa" Then
                        tgl5 = "0"
                    ElseIf dr.Item(6) = "Hadir" Then
                        tgl5 = "1"
                    End If
                    If dr.Item(7) Is DBNull.Value Or dr.Item(7) = "Sakit" Or dr.Item(7) = "Izin" Or dr.Item(7) = "Alfa" Then
                        tgl6 = "0"
                    ElseIf dr.Item(7) = "Hadir" Then
                        tgl6 = "1"
                    End If
                    If dr.Item(8) Is DBNull.Value Or dr.Item(8) = "Sakit" Or dr.Item(8) = "Izin" Or dr.Item(8) = "Alfa" Then
                        tgl7 = "0"
                    ElseIf dr.Item(8) = "Hadir" Then
                        tgl7 = "1"
                    End If
                    If dr.Item(9) Is DBNull.Value Or dr.Item(9) = "Sakit" Or dr.Item(9) = "Izin" Or dr.Item(9) = "Alfa" Then
                        tgl8 = "0"
                    ElseIf dr.Item(9) = "Hadir" Then
                        tgl8 = "1"
                    End If
                    If dr.Item(10) Is DBNull.Value Or dr.Item(10) = "Sakit" Or dr.Item(10) = "Izin" Or dr.Item(10) = "Alfa" Then
                        tgl9 = "0"
                    ElseIf dr.Item(10) = "Hadir" Then
                        tgl9 = "1"
                    End If
                    If dr.Item(11) Is DBNull.Value Or dr.Item(11) = "Sakit" Or dr.Item(11) = "Izin" Or dr.Item(11) = "Alfa" Then
                        tgl10 = "0"
                    ElseIf dr.Item(11) = "Hadir" Then
                        tgl10 = "1"
                    End If
                    If dr.Item(12) Is DBNull.Value Or dr.Item(12) = "Sakit" Or dr.Item(12) = "Izin" Or dr.Item(12) = "Alfa" Then
                        tgl11 = "0"
                    ElseIf dr.Item(12) = "Hadir" Then
                        tgl11 = "1"
                    End If
                    If dr.Item(13) Is DBNull.Value Or dr.Item(13) = "Sakit" Or dr.Item(13) = "Izin" Or dr.Item(13) = "Alfa" Then
                        tgl12 = "0"
                    ElseIf dr.Item(13) = "Hadir" Then
                        tgl12 = "1"
                    End If
                    If dr.Item(14) Is DBNull.Value Or dr.Item(14) = "Sakit" Or dr.Item(14) = "Izin" Or dr.Item(14) = "Alfa" Then
                        tgl13 = "0"
                    ElseIf dr.Item(14) = "Hadir" Then
                        tgl13 = "1"
                    End If
                    If dr.Item(15) Is DBNull.Value Or dr.Item(15) = "Sakit" Or dr.Item(15) = "Izin" Or dr.Item(15) = "Alfa" Then
                        tgl14 = "0"
                    ElseIf dr.Item(15) = "Hadir" Then
                        tgl14 = "1"
                    End If
                    If dr.Item(16) Is DBNull.Value Or dr.Item(16) = "Sakit" Or dr.Item(16) = "Izin" Or dr.Item(16) = "Alfa" Then
                        tgl15 = "0"
                    ElseIf dr.Item(16) = "Hadir" Then
                        tgl15 = "1"
                    End If
                    If dr.Item(17) Is DBNull.Value Or dr.Item(17) = "Sakit" Or dr.Item(17) = "Izin" Or dr.Item(17) = "Alfa" Then
                        tgl16 = "0"
                    ElseIf dr.Item(17) = "Hadir" Then
                        tgl16 = "1"
                    End If
                    If dr.Item(18) Is DBNull.Value Or dr.Item(18) = "Sakit" Or dr.Item(18) = "Izin" Or dr.Item(18) = "Alfa" Then
                        tgl17 = "0"
                    ElseIf dr.Item(18) = "Hadir" Then
                        tgl17 = "1"
                    End If
                    If dr.Item(19) Is DBNull.Value Or dr.Item(19) = "Sakit" Or dr.Item(19) = "Izin" Or dr.Item(19) = "Alfa" Then
                        tgl18 = "0"
                    ElseIf dr.Item(19) = "Hadir" Then
                        tgl18 = "1"
                    End If
                    If dr.Item(20) Is DBNull.Value Or dr.Item(20) = "Sakit" Or dr.Item(20) = "Izin" Or dr.Item(20) = "Alfa" Then
                        tgl19 = "0"
                    ElseIf dr.Item(20) = "Hadir" Then
                        tgl19 = "1"
                    End If
                    If dr.Item(21) Is DBNull.Value Or dr.Item(21) = "Sakit" Or dr.Item(21) = "Izin" Or dr.Item(21) = "Alfa" Then
                        tgl20 = "0"
                    ElseIf dr.Item(21) = "Hadir" Then
                        tgl20 = "1"
                    End If
                    If dr.Item(22) Is DBNull.Value Or dr.Item(22) = "Sakit" Or dr.Item(22) = "Izin" Or dr.Item(22) = "Alfa" Then
                        tgl21 = "0"
                    ElseIf dr.Item(22) = "Hadir" Then
                        tgl21 = "1"
                    End If
                    If dr.Item(23) Is DBNull.Value Or dr.Item(23) = "Sakit" Or dr.Item(23) = "Izin" Or dr.Item(23) = "Alfa" Then
                        tgl22 = "0"
                    ElseIf dr.Item(23) = "Hadir" Then
                        tgl22 = "1"
                    End If
                    If dr.Item(24) Is DBNull.Value Or dr.Item(24) = "Sakit" Or dr.Item(24) = "Izin" Or dr.Item(24) = "Alfa" Then
                        tgl23 = "0"
                    ElseIf dr.Item(24) = "Hadir" Then
                        tgl23 = "1"
                    End If
                    If dr.Item(25) Is DBNull.Value Or dr.Item(25) = "Sakit" Or dr.Item(25) = "Izin" Or dr.Item(25) = "Alfa" Then
                        tgl24 = "0"
                    ElseIf dr.Item(25) = "Hadir" Then
                        tgl24 = "1"
                    End If
                    If dr.Item(26) Is DBNull.Value Or dr.Item(26) = "Sakit" Or dr.Item(26) = "Izin" Or dr.Item(26) = "Alfa" Then
                        tgl25 = "0"
                    ElseIf dr.Item(26) = "Hadir" Then
                        tgl25 = "1"
                    End If
                    If dr.Item(27) Is DBNull.Value Or dr.Item(27) = "Sakit" Or dr.Item(27) = "Izin" Or dr.Item(27) = "Alfa" Then
                        tgl26 = "0"
                    ElseIf dr.Item(27) = "Hadir" Then
                        tgl26 = "1"
                    End If
                    If dr.Item(28) Is DBNull.Value Or dr.Item(28) = "Sakit" Or dr.Item(28) = "Izin" Or dr.Item(28) = "Alfa" Then
                        tgl27 = "0"
                    ElseIf dr.Item(28) = "Hadir" Then
                        tgl27 = "1"
                    End If
                    If dr.Item(29) Is DBNull.Value Or dr.Item(29) = "Sakit" Or dr.Item(29) = "Izin" Or dr.Item(29) = "Alfa" Then
                        tgl28 = "0"
                    ElseIf dr.Item(29) = "Hadir" Then
                        tgl28 = "1"
                    End If
                    If dr.Item(30) Is DBNull.Value Or dr.Item(30) = "Sakit" Or dr.Item(30) = "Izin" Or dr.Item(30) = "Alfa" Then
                        tgl29 = "0"
                    ElseIf dr.Item(30) = "Hadir" Then
                        tgl29 = "1"
                    End If
                    If dr.Item(31) Is DBNull.Value Or dr.Item(31) = "Sakit" Or dr.Item(31) = "Izin" Or dr.Item(31) = "Alfa" Then
                        tgl30 = "0"
                    ElseIf dr.Item(31) = "Hadir" Then
                        tgl30 = "1"
                    End If
                    If dr.Item(32) Is DBNull.Value Or dr.Item(32) = "Sakit" Or dr.Item(32) = "Izin" Or dr.Item(32) = "Alfa" Then
                        tgl31 = "0"
                    ElseIf dr.Item(32) = "Hadir" Then
                        tgl31 = "1"
                    End If
                    hasil = tgl1 + tgl2 + tgl3 + tgl4 + tgl5 + tgl6 + tgl7 + tgl8 + tgl9 + tgl10 + tgl11 + tgl12 + tgl13 + tgl14 + tgl15 + tgl16 + tgl17 + tgl18 + tgl19 + tgl20 + tgl21 + tgl22 + tgl23 + tgl24 + tgl25 + tgl26 + tgl27 + tgl28 + tgl29 + tgl30 + tgl31
                    Label15.Text = hasil
                End If

                If e.KeyChar = Chr(13) Then
                    Call carikode2()
                    If dr.Item(2) Is DBNull.Value Or dr.Item(2) = "Hadir" Or dr.Item(2) = "Izin" Or dr.Item(2) = "Alfa" Then
                        tgl1 = "0"
                    ElseIf dr.Item(2) = "Sakit" Then
                        tgl1 = "1"
                    End If
                    If dr.Item(3) Is DBNull.Value Or dr.Item(3) = "Hadir" Or dr.Item(3) = "Izin" Or dr.Item(3) = "Alfa" Then
                        tgl2 = "0"
                    ElseIf dr.Item(3) = "Sakit" Then
                        tgl2 = "1"
                    End If
                    If dr.Item(4) Is DBNull.Value Or dr.Item(4) = "Hadir" Or dr.Item(4) = "Izin" Or dr.Item(4) = "Alfa" Then
                        tgl3 = "0"
                    ElseIf dr.Item(4) = "Sakit" Then
                        tgl3 = "1"
                    End If
                    If dr.Item(5) Is DBNull.Value Or dr.Item(5) = "Hadir" Or dr.Item(5) = "Izin" Or dr.Item(5) = "Alfa" Then
                        tgl4 = "0"
                    ElseIf dr.Item(5) = "Sakit" Then
                        tgl4 = "1"
                    End If
                    If dr.Item(6) Is DBNull.Value Or dr.Item(6) = "Hadir" Or dr.Item(6) = "Izin" Or dr.Item(6) = "Alfa" Then
                        tgl5 = "0"
                    ElseIf dr.Item(6) = "Sakit" Then
                        tgl5 = "1"
                    End If
                    If dr.Item(7) Is DBNull.Value Or dr.Item(7) = "Hadir" Or dr.Item(7) = "Izin" Or dr.Item(7) = "Alfa" Then
                        tgl6 = "0"
                    ElseIf dr.Item(7) = "Sakit" Then
                        tgl6 = "1"
                    End If
                    If dr.Item(8) Is DBNull.Value Or dr.Item(8) = "Hadir" Or dr.Item(8) = "Izin" Or dr.Item(8) = "Alfa" Then
                        tgl7 = "0"
                    ElseIf dr.Item(8) = "Sakit" Then
                        tgl7 = "1"
                    End If
                    If dr.Item(9) Is DBNull.Value Or dr.Item(9) = "Hadir" Or dr.Item(9) = "Izin" Or dr.Item(9) = "Alfa" Then
                        tgl8 = "0"
                    ElseIf dr.Item(9) = "Sakit" Then
                        tgl8 = "1"
                    End If
                    If dr.Item(10) Is DBNull.Value Or dr.Item(10) = "Hadir" Or dr.Item(10) = "Izin" Or dr.Item(10) = "Alfa" Then
                        tgl9 = "0"
                    ElseIf dr.Item(10) = "Sakit" Then
                        tgl9 = "1"
                    End If
                    If dr.Item(11) Is DBNull.Value Or dr.Item(11) = "Hadir" Or dr.Item(11) = "Izin" Or dr.Item(11) = "Alfa" Then
                        tgl10 = "0"
                    ElseIf dr.Item(11) = "Sakit" Then
                        tgl10 = "1"
                    End If
                    If dr.Item(12) Is DBNull.Value Or dr.Item(12) = "Hadir" Or dr.Item(12) = "Izin" Or dr.Item(12) = "Alfa" Then
                        tgl11 = "0"
                    ElseIf dr.Item(12) = "Sakit" Then
                        tgl11 = "1"
                    End If
                    If dr.Item(13) Is DBNull.Value Or dr.Item(13) = "Hadir" Or dr.Item(13) = "Izin" Or dr.Item(13) = "Alfa" Then
                        tgl12 = "0"
                    ElseIf dr.Item(13) = "Sakit" Then
                        tgl12 = "1"
                    End If
                    If dr.Item(14) Is DBNull.Value Or dr.Item(14) = "Hadir" Or dr.Item(14) = "Izin" Or dr.Item(14) = "Alfa" Then
                        tgl13 = "0"
                    ElseIf dr.Item(14) = "Sakit" Then
                        tgl13 = "1"
                    End If
                    If dr.Item(15) Is DBNull.Value Or dr.Item(15) = "Hadir" Or dr.Item(15) = "Izin" Or dr.Item(15) = "Alfa" Then
                        tgl14 = "0"
                    ElseIf dr.Item(15) = "Sakit" Then
                        tgl14 = "1"
                    End If
                    If dr.Item(16) Is DBNull.Value Or dr.Item(16) = "Hadir" Or dr.Item(16) = "Izin" Or dr.Item(16) = "Alfa" Then
                        tgl15 = "0"
                    ElseIf dr.Item(16) = "Sakit" Then
                        tgl15 = "1"
                    End If
                    If dr.Item(17) Is DBNull.Value Or dr.Item(17) = "Hadir" Or dr.Item(17) = "Izin" Or dr.Item(17) = "Alfa" Then
                        tgl16 = "0"
                    ElseIf dr.Item(17) = "Sakit" Then
                        tgl16 = "1"
                    End If
                    If dr.Item(18) Is DBNull.Value Or dr.Item(18) = "Hadir" Or dr.Item(18) = "Izin" Or dr.Item(18) = "Alfa" Then
                        tgl17 = "0"
                    ElseIf dr.Item(18) = "Sakit" Then
                        tgl17 = "1"
                    End If
                    If dr.Item(19) Is DBNull.Value Or dr.Item(19) = "Hadir" Or dr.Item(19) = "Izin" Or dr.Item(19) = "Alfa" Then
                        tgl18 = "0"
                    ElseIf dr.Item(19) = "Sakit" Then
                        tgl18 = "1"
                    End If
                    If dr.Item(20) Is DBNull.Value Or dr.Item(20) = "Hadir" Or dr.Item(20) = "Izin" Or dr.Item(20) = "Alfa" Then
                        tgl19 = "0"
                    ElseIf dr.Item(20) = "Sakit" Then
                        tgl19 = "1"
                    End If
                    If dr.Item(21) Is DBNull.Value Or dr.Item(21) = "Hadir" Or dr.Item(21) = "Izin" Or dr.Item(21) = "Alfa" Then
                        tgl20 = "0"
                    ElseIf dr.Item(21) = "Sakit" Then
                        tgl20 = "1"
                    End If
                    If dr.Item(22) Is DBNull.Value Or dr.Item(22) = "Hadir" Or dr.Item(22) = "Izin" Or dr.Item(22) = "Alfa" Then
                        tgl21 = "0"
                    ElseIf dr.Item(22) = "Sakit" Then
                        tgl21 = "1"
                    End If
                    If dr.Item(23) Is DBNull.Value Or dr.Item(23) = "Hadir" Or dr.Item(23) = "Izin" Or dr.Item(23) = "Alfa" Then
                        tgl22 = "0"
                    ElseIf dr.Item(23) = "Sakit" Then
                        tgl22 = "1"
                    End If
                    If dr.Item(24) Is DBNull.Value Or dr.Item(24) = "Hadir" Or dr.Item(24) = "Izin" Or dr.Item(24) = "Alfa" Then
                        tgl23 = "0"
                    ElseIf dr.Item(24) = "Sakit" Then
                        tgl23 = "1"
                    End If
                    If dr.Item(25) Is DBNull.Value Or dr.Item(25) = "Hadir" Or dr.Item(25) = "Izin" Or dr.Item(25) = "Alfa" Then
                        tgl24 = "0"
                    ElseIf dr.Item(25) = "Sakit" Then
                        tgl24 = "1"
                    End If
                    If dr.Item(26) Is DBNull.Value Or dr.Item(26) = "Hadir" Or dr.Item(26) = "Izin" Or dr.Item(26) = "Alfa" Then
                        tgl25 = "0"
                    ElseIf dr.Item(26) = "Sakit" Then
                        tgl25 = "1"
                    End If
                    If dr.Item(27) Is DBNull.Value Or dr.Item(27) = "Hadir" Or dr.Item(27) = "Izin" Or dr.Item(27) = "Alfa" Then
                        tgl26 = "0"
                    ElseIf dr.Item(27) = "Sakit" Then
                        tgl26 = "1"
                    End If
                    If dr.Item(28) Is DBNull.Value Or dr.Item(28) = "Hadir" Or dr.Item(28) = "Izin" Or dr.Item(28) = "Alfa" Then
                        tgl27 = "0"
                    ElseIf dr.Item(28) = "Sakit" Then
                        tgl27 = "1"
                    End If
                    If dr.Item(29) Is DBNull.Value Or dr.Item(29) = "Hadir" Or dr.Item(29) = "Izin" Or dr.Item(29) = "Alfa" Then
                        tgl28 = "0"
                    ElseIf dr.Item(29) = "Sakit" Then
                        tgl28 = "1"
                    End If
                    If dr.Item(30) Is DBNull.Value Or dr.Item(30) = "Hadir" Or dr.Item(30) = "Izin" Or dr.Item(30) = "Alfa" Then
                        tgl29 = "0"
                    ElseIf dr.Item(30) = "Sakit" Then
                        tgl29 = "1"
                    End If
                    If dr.Item(31) Is DBNull.Value Or dr.Item(31) = "Hadir" Or dr.Item(31) = "Izin" Or dr.Item(31) = "Alfa" Then
                        tgl30 = "0"
                    ElseIf dr.Item(31) = "Sakit" Then
                        tgl30 = "1"
                    End If
                    If dr.Item(32) Is DBNull.Value Or dr.Item(32) = "Hadir" Or dr.Item(32) = "Izin" Or dr.Item(32) = "Alfa" Then
                        tgl31 = "0"
                    ElseIf dr.Item(32) = "Sakit" Then
                        tgl31 = "1"
                    End If
                    hasil = tgl1 + tgl2 + tgl3 + tgl4 + tgl5 + tgl6 + tgl7 + tgl8 + tgl9 + tgl10 + tgl11 + tgl12 + tgl13 + tgl14 + tgl15 + tgl16 + tgl17 + tgl18 + tgl19 + tgl20 + tgl21 + tgl22 + tgl23 + tgl24 + tgl25 + tgl26 + tgl27 + tgl28 + tgl29 + tgl30 + tgl31
                    Label9.Text = hasil
                End If

                If e.KeyChar = Chr(13) Then
                    Call carikode2()
                    If dr.Item(2) Is DBNull.Value Or dr.Item(2) = "Hadir" Or dr.Item(2) = "Sakit" Or dr.Item(2) = "Alfa" Then
                        tgl1 = "0"
                    ElseIf dr.Item(2) = "Izin" Then
                        tgl1 = "1"
                    End If
                    If dr.Item(3) Is DBNull.Value Or dr.Item(3) = "Hadir" Or dr.Item(3) = "Sakit" Or dr.Item(3) = "Alfa" Then
                        tgl2 = "0"
                    ElseIf dr.Item(3) = "Izin" Then
                        tgl2 = "1"
                    End If
                    If dr.Item(4) Is DBNull.Value Or dr.Item(4) = "Hadir" Or dr.Item(4) = "Sakit" Or dr.Item(4) = "Alfa" Then
                        tgl3 = "0"
                    ElseIf dr.Item(4) = "Izin" Then
                        tgl3 = "1"
                    End If
                    If dr.Item(5) Is DBNull.Value Or dr.Item(5) = "Hadir" Or dr.Item(5) = "Sakit" Or dr.Item(5) = "Alfa" Then
                        tgl4 = "0"
                    ElseIf dr.Item(5) = "Izin" Then
                        tgl4 = "1"
                    End If
                    If dr.Item(6) Is DBNull.Value Or dr.Item(6) = "Hadir" Or dr.Item(6) = "Sakit" Or dr.Item(6) = "Alfa" Then
                        tgl5 = "0"
                    ElseIf dr.Item(6) = "Izin" Then
                        tgl5 = "1"
                    End If
                    If dr.Item(7) Is DBNull.Value Or dr.Item(7) = "Hadir" Or dr.Item(7) = "Sakit" Or dr.Item(7) = "Alfa" Then
                        tgl6 = "0"
                    ElseIf dr.Item(7) = "Izin" Then
                        tgl6 = "1"
                    End If
                    If dr.Item(8) Is DBNull.Value Or dr.Item(8) = "Hadir" Or dr.Item(8) = "Sakit" Or dr.Item(8) = "Alfa" Then
                        tgl7 = "0"
                    ElseIf dr.Item(8) = "Izin" Then
                        tgl7 = "1"
                    End If
                    If dr.Item(9) Is DBNull.Value Or dr.Item(9) = "Hadir" Or dr.Item(9) = "Sakit" Or dr.Item(9) = "Alfa" Then
                        tgl8 = "0"
                    ElseIf dr.Item(9) = "Izin" Then
                        tgl8 = "1"
                    End If
                    If dr.Item(10) Is DBNull.Value Or dr.Item(10) = "Hadir" Or dr.Item(10) = "Sakit" Or dr.Item(10) = "Alfa" Then
                        tgl9 = "0"
                    ElseIf dr.Item(10) = "Izin" Then
                        tgl9 = "1"
                    End If
                    If dr.Item(11) Is DBNull.Value Or dr.Item(11) = "Hadir" Or dr.Item(11) = "Sakit" Or dr.Item(11) = "Alfa" Then
                        tgl10 = "0"
                    ElseIf dr.Item(11) = "Izin" Then
                        tgl10 = "1"
                    End If
                    If dr.Item(12) Is DBNull.Value Or dr.Item(12) = "Hadir" Or dr.Item(12) = "Sakit" Or dr.Item(12) = "Alfa" Then
                        tgl11 = "0"
                    ElseIf dr.Item(12) = "Izin" Then
                        tgl11 = "1"
                    End If
                    If dr.Item(13) Is DBNull.Value Or dr.Item(13) = "Hadir" Or dr.Item(13) = "Sakit" Or dr.Item(13) = "Alfa" Then
                        tgl12 = "0"
                    ElseIf dr.Item(13) = "Izin" Then
                        tgl12 = "1"
                    End If
                    If dr.Item(14) Is DBNull.Value Or dr.Item(14) = "Hadir" Or dr.Item(14) = "Sakit" Or dr.Item(14) = "Alfa" Then
                        tgl13 = "0"
                    ElseIf dr.Item(14) = "Izin" Then
                        tgl13 = "1"
                    End If
                    If dr.Item(15) Is DBNull.Value Or dr.Item(15) = "Hadir" Or dr.Item(15) = "Sakit" Or dr.Item(15) = "Alfa" Then
                        tgl14 = "0"
                    ElseIf dr.Item(15) = "Izin" Then
                        tgl14 = "1"
                    End If
                    If dr.Item(16) Is DBNull.Value Or dr.Item(16) = "Hadir" Or dr.Item(16) = "Sakit" Or dr.Item(16) = "Alfa" Then
                        tgl15 = "0"
                    ElseIf dr.Item(16) = "Izin" Then
                        tgl15 = "1"
                    End If
                    If dr.Item(17) Is DBNull.Value Or dr.Item(17) = "Hadir" Or dr.Item(17) = "Sakit" Or dr.Item(17) = "Alfa" Then
                        tgl16 = "0"
                    ElseIf dr.Item(17) = "Izin" Then
                        tgl16 = "1"
                    End If
                    If dr.Item(18) Is DBNull.Value Or dr.Item(18) = "Hadir" Or dr.Item(18) = "Sakit" Or dr.Item(18) = "Alfa" Then
                        tgl17 = "0"
                    ElseIf dr.Item(18) = "Izin" Then
                        tgl17 = "1"
                    End If
                    If dr.Item(19) Is DBNull.Value Or dr.Item(19) = "Hadir" Or dr.Item(19) = "Sakit" Or dr.Item(19) = "Alfa" Then
                        tgl18 = "0"
                    ElseIf dr.Item(19) = "Izin" Then
                        tgl18 = "1"
                    End If
                    If dr.Item(20) Is DBNull.Value Or dr.Item(20) = "Hadir" Or dr.Item(20) = "Sakit" Or dr.Item(20) = "Alfa" Then
                        tgl19 = "0"
                    ElseIf dr.Item(20) = "Izin" Then
                        tgl19 = "1"
                    End If
                    If dr.Item(21) Is DBNull.Value Or dr.Item(21) = "Hadir" Or dr.Item(21) = "Sakit" Or dr.Item(21) = "Alfa" Then
                        tgl20 = "0"
                    ElseIf dr.Item(21) = "Izin" Then
                        tgl20 = "1"
                    End If
                    If dr.Item(22) Is DBNull.Value Or dr.Item(22) = "Hadir" Or dr.Item(22) = "Sakit" Or dr.Item(22) = "Alfa" Then
                        tgl21 = "0"
                    ElseIf dr.Item(22) = "Izin" Then
                        tgl21 = "1"
                    End If
                    If dr.Item(23) Is DBNull.Value Or dr.Item(23) = "Hadir" Or dr.Item(23) = "Sakit" Or dr.Item(23) = "Alfa" Then
                        tgl22 = "0"
                    ElseIf dr.Item(23) = "Izin" Then
                        tgl22 = "1"
                    End If
                    If dr.Item(24) Is DBNull.Value Or dr.Item(24) = "Hadir" Or dr.Item(24) = "Sakit" Or dr.Item(24) = "Alfa" Then
                        tgl23 = "0"
                    ElseIf dr.Item(24) = "Izin" Then
                        tgl23 = "1"
                    End If
                    If dr.Item(25) Is DBNull.Value Or dr.Item(25) = "Hadir" Or dr.Item(25) = "Sakit" Or dr.Item(25) = "Alfa" Then
                        tgl24 = "0"
                    ElseIf dr.Item(25) = "Izin" Then
                        tgl24 = "1"
                    End If
                    If dr.Item(26) Is DBNull.Value Or dr.Item(26) = "Hadir" Or dr.Item(26) = "Sakit" Or dr.Item(26) = "Alfa" Then
                        tgl25 = "0"
                    ElseIf dr.Item(26) = "Izin" Then
                        tgl25 = "1"
                    End If
                    If dr.Item(27) Is DBNull.Value Or dr.Item(27) = "Hadir" Or dr.Item(27) = "Sakit" Or dr.Item(27) = "Alfa" Then
                        tgl26 = "0"
                    ElseIf dr.Item(27) = "Izin" Then
                        tgl26 = "1"
                    End If
                    If dr.Item(28) Is DBNull.Value Or dr.Item(28) = "Hadir" Or dr.Item(28) = "Sakit" Or dr.Item(28) = "Alfa" Then
                        tgl27 = "0"
                    ElseIf dr.Item(28) = "Izin" Then
                        tgl27 = "1"
                    End If
                    If dr.Item(29) Is DBNull.Value Or dr.Item(29) = "Hadir" Or dr.Item(29) = "Sakit" Or dr.Item(29) = "Alfa" Then
                        tgl28 = "0"
                    ElseIf dr.Item(29) = "Izin" Then
                        tgl28 = "1"
                    End If
                    If dr.Item(30) Is DBNull.Value Or dr.Item(30) = "Hadir" Or dr.Item(30) = "Sakit" Or dr.Item(30) = "Alfa" Then
                        tgl29 = "0"
                    ElseIf dr.Item(30) = "Izin" Then
                        tgl29 = "1"
                    End If
                    If dr.Item(31) Is DBNull.Value Or dr.Item(31) = "Hadir" Or dr.Item(31) = "Sakit" Or dr.Item(31) = "Alfa" Then
                        tgl30 = "0"
                    ElseIf dr.Item(31) = "Izin" Then
                        tgl30 = "1"
                    End If
                    If dr.Item(32) Is DBNull.Value Or dr.Item(32) = "Hadir" Or dr.Item(32) = "Sakit" Or dr.Item(32) = "Alfa" Then
                        tgl31 = "0"
                    ElseIf dr.Item(32) = "Izin" Then
                        tgl31 = "1"
                    End If
                    hasil = tgl1 + tgl2 + tgl3 + tgl4 + tgl5 + tgl6 + tgl7 + tgl8 + tgl9 + tgl10 + tgl11 + tgl12 + tgl13 + tgl14 + tgl15 + tgl16 + tgl17 + tgl18 + tgl19 + tgl20 + tgl21 + tgl22 + tgl23 + tgl24 + tgl25 + tgl26 + tgl27 + tgl28 + tgl29 + tgl30 + tgl31
                    Label11.Text = hasil
                End If


                If e.KeyChar = Chr(13) Then
                    Call carikode2()
                    If dr.Item(2) Is DBNull.Value Or dr.Item(2) = "Hadir" Or dr.Item(2) = "Sakit" Or dr.Item(2) = "Izin" Then
                        tgl1 = "0"
                    ElseIf dr.Item(2) = "Alfa" Then
                        tgl1 = "1"
                    End If
                    If dr.Item(3) Is DBNull.Value Or dr.Item(3) = "Hadir" Or dr.Item(3) = "Sakit" Or dr.Item(3) = "Izin" Then
                        tgl2 = "0"
                    ElseIf dr.Item(3) = "Alfa" Then
                        tgl2 = "1"
                    End If
                    If dr.Item(4) Is DBNull.Value Or dr.Item(4) = "Hadir" Or dr.Item(4) = "Sakit" Or dr.Item(4) = "Izin" Then
                        tgl3 = "0"
                    ElseIf dr.Item(4) = "Alfa" Then
                        tgl3 = "1"
                    End If
                    If dr.Item(5) Is DBNull.Value Or dr.Item(5) = "Hadir" Or dr.Item(5) = "Sakit" Or dr.Item(5) = "Izin" Then
                        tgl4 = "0"
                    ElseIf dr.Item(5) = "Alfa" Then
                        tgl4 = "1"
                    End If
                    If dr.Item(6) Is DBNull.Value Or dr.Item(6) = "Hadir" Or dr.Item(6) = "Sakit" Or dr.Item(6) = "Izin" Then
                        tgl5 = "0"
                    ElseIf dr.Item(6) = "Alfa" Then
                        tgl5 = "1"
                    End If
                    If dr.Item(7) Is DBNull.Value Or dr.Item(7) = "Hadir" Or dr.Item(7) = "Sakit" Or dr.Item(7) = "Izin" Then
                        tgl6 = "0"
                    ElseIf dr.Item(7) = "Alfa" Then
                        tgl6 = "1"
                    End If
                    If dr.Item(8) Is DBNull.Value Or dr.Item(8) = "Hadir" Or dr.Item(8) = "Sakit" Or dr.Item(8) = "Izin" Then
                        tgl7 = "0"
                    ElseIf dr.Item(8) = "Alfa" Then
                        tgl7 = "1"
                    End If
                    If dr.Item(9) Is DBNull.Value Or dr.Item(9) = "Hadir" Or dr.Item(9) = "Sakit" Or dr.Item(9) = "Izin" Then
                        tgl8 = "0"
                    ElseIf dr.Item(9) = "Alfa" Then
                        tgl8 = "1"
                    End If
                    If dr.Item(10) Is DBNull.Value Or dr.Item(10) = "Hadir" Or dr.Item(10) = "Sakit" Or dr.Item(10) = "Izin" Then
                        tgl9 = "0"
                    ElseIf dr.Item(10) = "Alfa" Then
                        tgl9 = "1"
                    End If
                    If dr.Item(11) Is DBNull.Value Or dr.Item(11) = "Hadir" Or dr.Item(11) = "Sakit" Or dr.Item(11) = "Izin" Then
                        tgl10 = "0"
                    ElseIf dr.Item(11) = "Alfa" Then
                        tgl10 = "1"
                    End If
                    If dr.Item(12) Is DBNull.Value Or dr.Item(12) = "Hadir" Or dr.Item(12) = "Sakit" Or dr.Item(12) = "Izin" Then
                        tgl11 = "0"
                    ElseIf dr.Item(12) = "Alfa" Then
                        tgl11 = "1"
                    End If
                    If dr.Item(13) Is DBNull.Value Or dr.Item(13) = "Hadir" Or dr.Item(13) = "Sakit" Or dr.Item(13) = "Izin" Then
                        tgl12 = "0"
                    ElseIf dr.Item(13) = "Alfa" Then
                        tgl12 = "1"
                    End If
                    If dr.Item(14) Is DBNull.Value Or dr.Item(14) = "Hadir" Or dr.Item(14) = "Sakit" Or dr.Item(14) = "Izin" Then
                        tgl13 = "0"
                    ElseIf dr.Item(14) = "Alfa" Then
                        tgl13 = "1"
                    End If
                    If dr.Item(15) Is DBNull.Value Or dr.Item(15) = "Hadir" Or dr.Item(15) = "Sakit" Or dr.Item(15) = "Izin" Then
                        tgl14 = "0"
                    ElseIf dr.Item(15) = "Alfa" Then
                        tgl14 = "1"
                    End If
                    If dr.Item(16) Is DBNull.Value Or dr.Item(16) = "Hadir" Or dr.Item(16) = "Sakit" Or dr.Item(16) = "Izin" Then
                        tgl15 = "0"
                    ElseIf dr.Item(16) = "Alfa" Then
                        tgl15 = "1"
                    End If
                    If dr.Item(17) Is DBNull.Value Or dr.Item(17) = "Hadir" Or dr.Item(17) = "Sakit" Or dr.Item(17) = "Izin" Then
                        tgl16 = "0"
                    ElseIf dr.Item(17) = "Alfa" Then
                        tgl16 = "1"
                    End If
                    If dr.Item(18) Is DBNull.Value Or dr.Item(18) = "Hadir" Or dr.Item(18) = "Sakit" Or dr.Item(18) = "Izin" Then
                        tgl17 = "0"
                    ElseIf dr.Item(18) = "Alfa" Then
                        tgl17 = "1"
                    End If
                    If dr.Item(19) Is DBNull.Value Or dr.Item(19) = "Hadir" Or dr.Item(19) = "Sakit" Or dr.Item(19) = "Izin" Then
                        tgl18 = "0"
                    ElseIf dr.Item(19) = "Alfa" Then
                        tgl18 = "1"
                    End If
                    If dr.Item(20) Is DBNull.Value Or dr.Item(20) = "Hadir" Or dr.Item(20) = "Sakit" Or dr.Item(20) = "Izin" Then
                        tgl19 = "0"
                    ElseIf dr.Item(20) = "Alfa" Then
                        tgl19 = "1"
                    End If
                    If dr.Item(21) Is DBNull.Value Or dr.Item(21) = "Hadir" Or dr.Item(21) = "Sakit" Or dr.Item(21) = "Izin" Then
                        tgl20 = "0"
                    ElseIf dr.Item(21) = "Alfa" Then
                        tgl20 = "1"
                    End If
                    If dr.Item(22) Is DBNull.Value Or dr.Item(22) = "Hadir" Or dr.Item(22) = "Sakit" Or dr.Item(22) = "Izin" Then
                        tgl21 = "0"
                    ElseIf dr.Item(22) = "Alfa" Then
                        tgl21 = "1"
                    End If
                    If dr.Item(23) Is DBNull.Value Or dr.Item(23) = "Hadir" Or dr.Item(23) = "Sakit" Or dr.Item(23) = "Izin" Then
                        tgl22 = "0"
                    ElseIf dr.Item(23) = "Alfa" Then
                        tgl22 = "1"
                    End If
                    If dr.Item(24) Is DBNull.Value Or dr.Item(24) = "Hadir" Or dr.Item(24) = "Sakit" Or dr.Item(24) = "Izin" Then
                        tgl23 = "0"
                    ElseIf dr.Item(24) = "Alfa" Then
                        tgl23 = "1"
                    End If
                    If dr.Item(25) Is DBNull.Value Or dr.Item(25) = "Hadir" Or dr.Item(25) = "Sakit" Or dr.Item(25) = "Izin" Then
                        tgl24 = "0"
                    ElseIf dr.Item(25) = "Alfa" Then
                        tgl24 = "1"
                    End If
                    If dr.Item(26) Is DBNull.Value Or dr.Item(26) = "Hadir" Or dr.Item(26) = "Sakit" Or dr.Item(26) = "Izin" Then
                        tgl25 = "0"
                    ElseIf dr.Item(26) = "Alfa" Then
                        tgl25 = "1"
                    End If
                    If dr.Item(27) Is DBNull.Value Or dr.Item(27) = "Hadir" Or dr.Item(27) = "Sakit" Or dr.Item(27) = "Izin" Then
                        tgl26 = "0"
                    ElseIf dr.Item(27) = "Alfa" Then
                        tgl26 = "1"
                    End If
                    If dr.Item(28) Is DBNull.Value Or dr.Item(28) = "Hadir" Or dr.Item(28) = "Sakit" Or dr.Item(28) = "Izin" Then
                        tgl27 = "0"
                    ElseIf dr.Item(28) = "Alfa" Then
                        tgl27 = "1"
                    End If
                    If dr.Item(29) Is DBNull.Value Or dr.Item(29) = "Hadir" Or dr.Item(29) = "Sakit" Or dr.Item(29) = "Izin" Then
                        tgl28 = "0"
                    ElseIf dr.Item(29) = "Alfa" Then
                        tgl28 = "1"
                    End If
                    If dr.Item(30) Is DBNull.Value Or dr.Item(30) = "Hadir" Or dr.Item(30) = "Sakit" Or dr.Item(30) = "Izin" Then
                        tgl29 = "0"
                    ElseIf dr.Item(30) = "Alfa" Then
                        tgl29 = "1"
                    End If
                    If dr.Item(31) Is DBNull.Value Or dr.Item(31) = "Hadir" Or dr.Item(31) = "Sakit" Or dr.Item(31) = "Izin" Then
                        tgl30 = "0"
                    ElseIf dr.Item(31) = "Alfa" Then
                        tgl30 = "1"
                    End If
                    If dr.Item(32) Is DBNull.Value Or dr.Item(32) = "Hadir" Or dr.Item(32) = "Sakit" Or dr.Item(32) = "Izin" Then
                        tgl31 = "0"
                    ElseIf dr.Item(32) = "Alfa" Then
                        tgl31 = "1"
                    End If
                    hasil = tgl1 + tgl2 + tgl3 + tgl4 + tgl5 + tgl6 + tgl7 + tgl8 + tgl9 + tgl10 + tgl11 + tgl12 + tgl13 + tgl14 + tgl15 + tgl16 + tgl17 + tgl18 + tgl19 + tgl20 + tgl21 + tgl22 + tgl23 + tgl24 + tgl25 + tgl26 + tgl27 + tgl28 + tgl29 + tgl30 + tgl31
                    Label13.Text = hasil
                End If

            ElseIf Not dr.HasRows Then
                MsgBox("Tidak ada siswa yang ber NIS " + TextBox3.Text)
            End If
        End If
        If Not ((e.KeyChar >= "0" And e.KeyChar <= "9") Or e.KeyChar = vbBack) Then
            e.Handled = True

        End If
    End Sub

    Private Sub TextBox3_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox3.TextChanged

    End Sub

    Private Sub Button9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button9.Click

        If RadioButton1.Checked = True Then
            pilihan = RadioButton1.Text
        ElseIf RadioButton2.Checked = True Then
            pilihan = RadioButton2.Text
        ElseIf RadioButton3.Checked = True Then
            pilihan = RadioButton3.Text
        ElseIf RadioButton4.Checked = True Then
            pilihan = RadioButton4.Text

        End If
        If DateTimePicker1.Text = "01 Juli 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_Juli set  1 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()

        ElseIf DateTimePicker1.Text = "02 Juli 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_juli set  2 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "03 Juli 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_juli set  3 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "04 Juli 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_juli set  4 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "05 Juli 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_juli set  5 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "06 Juli 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_juli set  6 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "07 Juli 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_juli set  8 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "09 Juli 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_juli set  9 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "10 Juli 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_juli set  10 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "11 Juli 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_juli set  11 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "12 Juli 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_juli set  12 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "13 Juli 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_juli set  13 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "14 Juli 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_juli set  14 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "15 Juli 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_juli set  15 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "16 Juli 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_juli set  16 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "17 Juli 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_juli set  17 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "18 Juli 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_juli set  18 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "19 Juli 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_juli set  19 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "20 Juli 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_juli set  20 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "21 Juli 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_juli set  21 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "22 Juli 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_juli set  22 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "23 Juli 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_juli set  23 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "24 Juli 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_juli set  24 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "25 Juli 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_juli set  25 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "26 Juli 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_juli set 26 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "27 Juli 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_juli set  27 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "28 Juli 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_juli set  28 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "29 Juli 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_juli set  29 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "30 Juli 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_juli set  30 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "31 Juli 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_juli set  31 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()



        ElseIf DateTimePicker1.Text = "01 Agustus 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_agustus set  1 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "02 Agustus 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_agustus set  2 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "03 Agustus 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_agustus set  3 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "04 Agustus 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_agustus set  4 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "05 Agustus 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_agustus set  5 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "06 Agustus 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_agustus set  6 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "07 Agustus 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_agustus set  7 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "08 Agustus 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_agustus set  8 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "09 Agustus 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_agustus set  9 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "10 Agustus 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_agustus set  10 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "11 Agustus 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_agustus set  11 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "12 Agustus 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_agustus set  12 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "13 Agustus 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_agustus set  13 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "14 Agustus 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_agustus set  14 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "15 Agustus 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_agustus set  15 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "16 Agustus 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_agustus set  16 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "17 Agustus 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_agustus set  17 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "18 Agustus 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_agustus set  18 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "19 Agustus 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_agustus set  19 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "20 Agustus 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_agustus set  20 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "21 Agustus 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_agustus set  21 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "22 Agustus 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_agustus set  22 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "23 Agustus 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_agustus set  23 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "24 Agustus 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_agustus set  24 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "25 Agustus 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_agustus set  25 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "26 Agustus 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_agustus set  26 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "27 Agustus 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_agustus set  27 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "28 Agustus 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_agustus set  28 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "29 Agustus 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_agustus set  29 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "30 Agustus 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_agustus set  30 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "31 Agustus 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_agustus set  31 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()


        ElseIf DateTimePicker1.Text = "01 September 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_september set  1 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "02 September 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_september set  2 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "03 September 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_september set  3 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "04 September 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_september set  4 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "05 September 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_september set  5 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "06 September 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_september set  6 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "07 September 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_september set  7 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "08 September 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_september set  8 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "09 September 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_september set  9 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "10 September 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_september set  10 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "11 September 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_september set  11 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "12 September 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_september set  12 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "13 September 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_september set  13 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "14 September 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_september set  14 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "15 September 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_september set  15 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "16 September 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_september set  16 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "17 September 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_september set  17 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "18 September 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_september set  18 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "19 September 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_september set  19 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "20 September 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_september set  20 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "21 September 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_september set  21 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "22 September 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_september set  22 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "23 September 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_september set  23 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "24 September 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_september set  24 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "25 September 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_september set  25 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "26 September 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_september set  26 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "27 September 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_september set  27 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "28 September 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_september set  28 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "29 September 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_september set  29 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "30 September 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_september set  30 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()



        ElseIf DateTimePicker1.Text = "01 Oktober 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_oktober set  1 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "02 Oktober 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_oktober set  2 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "03 Oktober 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_oktober set  3 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "04 Oktober 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_oktober set  4 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "05 Oktober 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_oktober set  5 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "06 Oktober 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_oktober set  6 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "07 Oktober 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_oktober set  7 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "08 Oktober 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_oktober set  8 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "09 Oktober 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_oktober set  9 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "10 Oktober 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_oktober set  10 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "11 Oktober 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_oktober set  11 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "12 Oktober 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_oktober set  12 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "13 Oktober 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_oktober set  13 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "14 Oktober 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_oktober set  14 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "15 Oktober 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_oktober set  15 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "16 Oktober 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_oktober set  16 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "17 Oktober 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_oktober set  17 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "18 Oktober 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_oktober set  18 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "19 Oktober 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_oktober set  19 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "20 Oktober 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_oktober set  20 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "21 Oktober 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_oktober set  21 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "22 Oktober 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_oktober set  22 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "23 Oktober 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_oktober set  23 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "24 Oktober 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_oktober set  24 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "25 Oktober 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_oktober set  25 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "26 Oktober 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_oktober set  26 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "27 Oktober 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_oktober set  27 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "28 Oktober 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_oktober set  28 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "29 Oktober 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_oktober set  29 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "30 Oktober 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_oktober set  30 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "31 Oktober 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_oktober set  31 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()


        ElseIf DateTimePicker1.Text = "01 November 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_november set  1 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "02 November 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_november set  2 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "03 November 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_november set  3 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "04 November 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_november set  4 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "05 November 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_november set  5 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "06 November 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_november set  6 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "07 November 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_november set  7 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "08 November 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_november set  8 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "09 November 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_november set  9 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "10 November 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_november set  10 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "11 November 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_november set  11 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "12 November 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_november set  12 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "13 November 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_november set  13 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "14 November 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_november set  14 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "15 November 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_november set  15 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "16 November 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_november set  16 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "17 November 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_november set  17 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "18 November 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_november set  18 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "19 November 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_november set  19 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "20 November 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_november set  20 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "21 November 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_november set  21 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "22 November 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_november set  22 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "23 November 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_november set  23 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "24 November 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_november set  24 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "25 November 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_november set  25 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "26 November 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_november set  26 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "27 November 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_november set  27 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "28 November 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_november set  28 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "29 November 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_november set  29 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "30 November 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_november set  30 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()


        ElseIf DateTimePicker1.Text = "01 Desember 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_desember set  1 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "02 Desember 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_desember set  2 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "03 Desember 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_desember set  3 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "04 Desember 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_desember set  4 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "05 Desember 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_desember set  5 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "06 Desember 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_desember set  6 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "07 Desember 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_desember set  7 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "08 Desember 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_desember set  8 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "09 Desember 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_desember set  9 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "10 Desember 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_desember set  10 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "11 Desember 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_desember set  11 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "12 Desember 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_desember set  12 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "13 Desember 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_desember set  13 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "14 Desember 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_desember set  14 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "15 Desember 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_desember set  15 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "16 Desember 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_desember set  16 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "17 Desember 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_desember set  17 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "18 Desember 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_desember set  18 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "19 Desember 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_desember set  19 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "20 Desember 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_desember set  20 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "21 Desember 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_desember set  21 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "22 Desember 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_desember set  22 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "23 Desember 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_desember set  23 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "24 Desember 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_desember set  24 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "25 Desember 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_desember set  25 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "26 Desember 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_desember set  26 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "27 Desember 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_desember set  27 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "28 Desember 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_desember set  28 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "29 Desember 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_desember set  29 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "30 Desember 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_desember set  30 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "31 Desember 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_desember set  31 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()


        ElseIf DateTimePicker1.Text = "01 Januari 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_januari set  1 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "02 Januari 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_januari set  2 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "03 Januari 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_januari set  3 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "04 Januari 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_januari set  4 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "05 Januari 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_januari set  5 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "06 Januari 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_januari set  6 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "07 Januari 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_januari set  7 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "08 Januari 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_januari set  8 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "09 Januari 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_januari set  9 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "10 Januari 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_januari set  10 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "11 Januari 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_januari set  11 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "12 Januari 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_januari set  12 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "13 Januari 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_januari set  13 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "14 Januari 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_januari set  14 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "15 Januari 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_januari set  15 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "16 Januari 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_januari set  16 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "17 Januari 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_januari set  17 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "18 Januari 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_januari set  18 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "19 Januari 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_januari set  19 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "20 Januari 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_januari set  20 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "21 Januari 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_januari set  21 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "22 Januari 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_januari set  22 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "23 Januari 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_januari set  23 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "24 Januari 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_januari set  24 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "25 Januari 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_januari set  25 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "26 Januari 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_januari set  26 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "27 Januari 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_januari set  27 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "28 Januari 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_januari set  28 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "29 Januari 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_januari set  29 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "30 Januari 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_januari set  30 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "31 Januari 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_januari set  31 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()


        ElseIf DateTimePicker1.Text = "01 Februari 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_februari set  1 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "02 Februari 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_februari set  2 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "03 Februari 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_februari set  3 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "04 Februari 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_februari set  4 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "05 Februari 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_februari set  5 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "06 Februari 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_februari set  6 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "07 Februari 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_februari set  7 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "08 Februari 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_februari set  8 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "09 Februari 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_februari set  9 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "10 Februari 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_februari set  10 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "11 Februari 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_februari set  11 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "12 Februari 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_februari set  12 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "13 Februari 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_februari set  13 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "14 Februari 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_februari set  14 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "15 Februari 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_februari set  15 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "16 Februari 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_februari set  16 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "17 Februari 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_februari set  17 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "18 Februari 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_februari set  18 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "19 Februari 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_februari set  19 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "20 Februari 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_februari set  20 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "21 Februari 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_februari set  21 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "22 Februari 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_februari set  22 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "23 Februari 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_februari set  23 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "24 Februari 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_februari set  24 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "25 Februari 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_februari set  25 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "26 Februari 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_februari set  26 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "27 Februari 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_februari set  27 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "28 Februari 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_februari set  28 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()


        ElseIf DateTimePicker1.Text = "01 Maret 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_maret set  1 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "02 Maret 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_maret set  2 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "03 Maret 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_maret set  3 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "04 Maret 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_maret set  4 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "05 Maret 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_maret set  5 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "06 Maret 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_maret set  6 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "07 Maret 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_maret set  7 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "08 Maret 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_maret set  8 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "09 Maret 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_maret set  9 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "10 Maret 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_maret set  10 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "11 Maret 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_maret set  11 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "12 Maret 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_maret set  12 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "13 Maret 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_maret set  13 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "14 Maret 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_maret set  14 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "15 Maret 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_maret set  15 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "16 Maret 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_maret set  16 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "17 Maret 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_maret set  17 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "18 Maret 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_maret set  18 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "19 Maret 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_maret set  19 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "20 Maret 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_maret set  20 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "21 Maret 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_maret set  21 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "22 Maret 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_maret set  22 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "23 Maret 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_maret set  23 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "24 Maret 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_maret set  24 ='" & pilihan & ", Keterangan='" & TextBox5.Text & "'' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "25 Maret 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_maret set  25 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "26 Maret 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_maret set  26 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "27 Maret 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_maret set  27 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "28 Maret 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_maret set  28 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "29 Maret 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_maret set  29 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "30 Maret 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_maret set  30 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "31 Maret 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_maret set  31 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()


        ElseIf DateTimePicker1.Text = "01 April 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_april set  1 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "02 April 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_april set  2 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "03 April 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_april set  3 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "04 April 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_april set  4 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "05 April 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_april set  5 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "06 April 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_april set  6 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "07 April 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_april set  7 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "08 April 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_april set  8 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "09 April 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_april set  9 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "10 April 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_april set  10 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "11 April 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_april set  11 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "12 April 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_april set  12 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "13 April 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_april set  13 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "14 April 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_april set  14 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "15 April 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_april set  15 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "16 April 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_april set  16 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "17 April 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_april set  17 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "18 April 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_april set  18 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "19 April 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_april set  19 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "20 April 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_april set  20 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "21 April 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_april set  21 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "22 April 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_april set  22 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "23 April 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_april set  23 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "24 April 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_april set  24 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "25 April 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_april set  25 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "26 April 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_april set  26 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "27 April 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_april set  27 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "28 April 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_april set  28 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "29 April 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_april set  29 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "30 April 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_april set  30 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()


        ElseIf DateTimePicker1.Text = "01 Mei 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_mei set  1 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "02 Mei 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_mei set  2 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "03 Mei 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_mei set  3 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "04 Mei 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_mei set  4 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "05 Mei 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_mei set  5 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "06 Mei 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_mei set  6 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "07 Mei 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_mei set  7 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "08 Mei 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_mei set  8 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "09 Mei 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_mei set  9 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "10 Mei 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_mei set  10 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "11 Mei 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_mei set  11 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "12 Mei 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_mei set  12 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "13 Mei 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_mei set  13 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "14 Mei 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_mei set  14 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "15 Mei 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_mei set  15 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "16 Mei 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_mei set  16 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "17 Mei 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_mei set  17 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "18 Mei 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_mei set  18 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "19 Mei 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_mei set  19 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "20 Mei 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_mei set  20 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "21 Mei 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_mei set  21 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "22 Mei 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_mei set  22 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "23 Mei 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_mei set  23 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "24 Mei 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_mei set  24 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "25 Mei 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_mei set  25 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "26 Mei 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_mei set  26 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "27 Mei 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_mei set  27 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "28 Mei 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_mei set  28 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "29 Mei 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_mei set  29 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "30 Mei 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_mei set  30 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "31 Mei 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_mei set  31 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()


        ElseIf DateTimePicker1.Text = "01 Juni 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_juni set  1 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "02 Juni 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_juni set  2 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "03 Juni 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_juni set  3 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "04 Juni 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_juni set  4 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "05 Juni 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_juni set  5 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "06 Juni 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_juni set  6 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "07 Juni 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_juni set  7 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "08 Juni 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_juni set  8 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "09 Juni 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_juni set  9 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "10 Juni 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_juni set  10 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "11 Juni 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_juni set  11 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "12 Juni 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_juni set  12 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "13 Juni 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_juni set  13 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "14 Juni 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_juni set  14 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "15 Juni 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_juni set  15 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "16 Juni 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_juni set  16 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "17 Juni 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_juni set  17 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "18 Juni 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_juni set  18 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "19 Juni 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_juni set  19 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "20 Juni 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_juni set  20 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "21 Juni 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_juni set  21 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "22 Juni 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_juni set  22 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "23 Juni 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_juni set  23 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "24 Juni 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_juni set  24 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "25 Juni 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_juni set  25 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "26 Juni 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_juni set  26 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "27 Juni 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_juni set  27 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "28 Juni 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_juni set  28 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "29 Juni 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_juni set  29 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "30 Juni 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_juni set  30 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "31 Juni 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_juni set  31 ='" & pilihan & "', Keterangan='" & TextBox5.Text & "' where NIS ='" & TextBox3.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()


        End If
    End Sub

    Private Sub Button10_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button10.Click
        If RadioButton1.Checked = True Then
            pilihan = RadioButton1.Text
        ElseIf RadioButton1.Checked = True Then
            pilihan = RadioButton2.Text
        ElseIf RadioButton3.Checked = True Then
            pilihan = RadioButton3.Text
        ElseIf RadioButton4.Checked = True Then
            pilihan = RadioButton4.Text

        End If
        If DateTimePicker1.Text = "01 Juli 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_Juli set  1 ='" & Button10.Text & "' "
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "02 Juli 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_juli set  2 ='" & Button10.Text & "' "
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "03 Juli 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_juli set  3 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "04 Juli 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_juli set  4 ='" & Button10.Text & "' "
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "05 Juli 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_juli set  5 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "06 Juli 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_juli set  6 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "07 Juli 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_juli set  8 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "09 Juli 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_juli set  9 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "10 Juli 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_juli set  10 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "11 Juli 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_juli set  11 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "12 Juli 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_juli set  12 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "13 Juli 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_juli set  13 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "14 Juli 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_juli set  14 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "15 Juli 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_juli set  15 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "16 Juli 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_juli set  16 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "17 Juli 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_juli set  17 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "18 Juli 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_juli set  18 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "19 Juli 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_juli set  19 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "20 Juli 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_juli set  20 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "21 Juli 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_juli set  21 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "22 Juli 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_juli set  22 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "23 Juli 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_juli set  23 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "24 Juli 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_juli set  24 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "25 Juli 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_juli set  25 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "26 Juli 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_juli set 26 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "27 Juli 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_juli set  27 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "28 Juli 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_juli set  28 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "29 Juli 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_juli set  29 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "30 Juli 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_juli set  30 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "31 Juli 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_juli set  31 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()



        ElseIf DateTimePicker1.Text = "01 Agustus 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_agustus set  1 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "02 Agustus 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_agustus set  2 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "03 Agustus 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_agustus set  3 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "04 Agustus 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_agustus set  4 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "05 Agustus 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_agustus set  5 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "06 Agustus 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_agustus set  6 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "07 Agustus 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_agustus set  7 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "08 Agustus 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_agustus set  8 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "09 Agustus 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_agustus set  9 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "10 Agustus 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_agustus set  10 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "11 Agustus 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_agustus set  11 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "12 Agustus 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_agustus set  12 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "13 Agustus 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_agustus set  13 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "14 Agustus 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_agustus set  14 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "15 Agustus 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_agustus set  15 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "16 Agustus 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_agustus set  16 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "17 Agustus 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_agustus set  17 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "18 Agustus 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_agustus set  18 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "19 Agustus 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_agustus set  19 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "20 Agustus 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_agustus set  20 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "21 Agustus 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_agustus set  21 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "22 Agustus 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_agustus set  22 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "23 Agustus 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_agustus set  23 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "24 Agustus 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_agustus set  24 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "25 Agustus 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_agustus set  25 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "26 Agustus 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_agustus set  26 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "27 Agustus 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_agustus set  27 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "28 Agustus 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_agustus set  28 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "29 Agustus 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_agustus set  29 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "30 Agustus 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_agustus set  30 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "31 Agustus 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_agustus set  31 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()


        ElseIf DateTimePicker1.Text = "01 September 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_september set  1 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "02 September 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_september set  2 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "03 September 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_september set  3 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "04 September 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_september set  4 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "05 September 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_september set  5 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "06 September 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_september set  6 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "07 September 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_september set  7 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "08 September 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_september set  8 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "09 September 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_september set  9 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "10 September 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_september set  10 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "11 September 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_september set  11 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "12 September 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_september set  12 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "13 September 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_september set  13 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "14 September 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_september set  14 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "15 September 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_september set  15 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "16 September 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_september set  16 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "17 September 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_september set  17 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "18 September 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_september set  18 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "19 September 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_september set  19 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "20 September 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_september set  20 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "21 September 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_september set  21 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "22 September 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_september set  22 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "23 September 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_september set  23 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "24 September 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_september set  24 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "25 September 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_september set  25 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "26 September 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_september set  26 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "27 September 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_september set  27 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "28 September 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_september set  28 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "29 September 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_september set  29 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "30 September 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_september set  30 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()



        ElseIf DateTimePicker1.Text = "01 Oktober 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_oktober set  1 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "02 Oktober 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_oktober set  2 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "03 Oktober 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_oktober set  3 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "04 Oktober 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_oktober set  4 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "05 Oktober 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_oktober set  5 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "06 Oktober 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_oktober set  6 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "07 Oktober 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_oktober set  7 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "08 Oktober 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_oktober set  8 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "09 Oktober 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_oktober set  9 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "10 Oktober 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_oktober set  10 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "11 Oktober 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_oktober set  11 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "12 Oktober 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_oktober set  12 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "13 Oktober 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_oktober set  13 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "14 Oktober 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_oktober set  14 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "15 Oktober 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_oktober set  15 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "16 Oktober 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_oktober set  16 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "17 Oktober 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_oktober set  17 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "18 Oktober 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_oktober set  18 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "19 Oktober 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_oktober set  19 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "20 Oktober 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_oktober set  20 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "21 Oktober 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_oktober set  21 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "22 Oktober 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_oktober set  2 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "23 Oktober 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_oktober set  23 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "24 Oktober 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_oktober set  24 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "25 Oktober 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_oktober set  25 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "26 Oktober 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_oktober set  26 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "27 Oktober 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_oktober set  27 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "28 Oktober 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_oktober set  28 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "29 Oktober 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_oktober set  29 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "30 Oktober 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_oktober set  30 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "31 Oktober 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_oktober set  31 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()


        ElseIf DateTimePicker1.Text = "01 November 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_november set  1 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "02 November 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_november set  2 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "03 November 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_november set  3 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "04 November 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_november set  4 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "05 November 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_november set  5 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "06 November 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_november set  6 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "07 November 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_november set  7 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "08 November 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_november set  8 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "09 November 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_november set  9 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "10 November 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_november set  10 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "11 November 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_november set  11 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "12 November 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_november set  12 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "13 November 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_november set  13 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "14 November 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_november set  14 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "15 November 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_november set  15 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "16 November 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_november set  16 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "17 November 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_november set  17 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "18 November 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_november set  18 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "19 November 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_november set  19 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "20 November 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_november set  20 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "21 November 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_november set  21 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "22 November 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_november set  22 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "23 November 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_november set  23 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "24 November 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_november set  24 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "25 November 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_november set  25 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "26 November 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_november set  26 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "27 November 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_november set  27 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "28 November 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_november set  28 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "29 November 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_november set  29 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "30 November 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_november set  30 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()


        ElseIf DateTimePicker1.Text = "01 Desember 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_desember set  1 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "02 Desember 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_desember set  2 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "03 Desember 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_desember set  3 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "04 Desember 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_desember set  4 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "05 Desember 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_desember set  5 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "06 Desember 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_desember set  6 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "07 Desember 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_desember set  7 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "08 Desember 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_desember set  8 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "09 Desember 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_desember set  9 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "10 Desember 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_desember set  10 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "11 Desember 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_desember set  11 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "12 Desember 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_desember set  12 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "13 Desember 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_desember set  13 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "14 Desember 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_desember set  14 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "15 Desember 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_desember set  15 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "16 Desember 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_desember set  16 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "17 Desember 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_desember set  17 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "18 Desember 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_desember set  18 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "19 Desember 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_desember set  19 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "20 Desember 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_desember set  20 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "21 Desember 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_desember set  21 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "22 Desember 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_desember set  22 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "23 Desember 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_desember set  23 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "24 Desember 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_desember set  24 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "25 Desember 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_desember set  25 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "26 Desember 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_desember set  26 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "27 Desember 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_desember set  27 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "28 Desember 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_desember set  28 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "29 Desember 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_desember set  29 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "30 Desember 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_desember set  30 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "31 Desember 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_desember set  31 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()


        ElseIf DateTimePicker1.Text = "01 Januari 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_januari set  1 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "02 Januari 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_januari set  2 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "03 Januari 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_januari set  3 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "04 Januari 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_januari set  4 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "05 Januari 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_januari set  5 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "06 Januari 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_januari set  6 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "07 Januari 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_januari set  7 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "08 Januari 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_januari set  8 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "09 Januari 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_januari set  9 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "10 Januari 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_januari set  10 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "11 Januari 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_januari set  11 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "12 Januari 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_januari set  12 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "13 Januari 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_januari set  13 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "14 Januari 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_januari set  14 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "15 Januari 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_januari set  15 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "16 Januari 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_januari set  16 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "17 Januari 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_januari set  17 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "18 Januari 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_januari set  18 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "19 Januari 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_januari set  19 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "20 Januari 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_januari set  20 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "21 Januari 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_januari set  21 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "22 Januari 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_januari set  22 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "23 Januari 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_januari set  23 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "24 Januari 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_januari set  24 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "25 Januari 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_januari set  25 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "26 Januari 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_januari set  26 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "27 Januari 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_januari set  27 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "28 Januari 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_januari set  28 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "29 Januari 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_januari set  29 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "30 Januari 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_januari set  30 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "31 Januari 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_januari set  31 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()


        ElseIf DateTimePicker1.Text = "01 Februari 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_februari set  1 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "02 Februari 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_februari set  2 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "03 Februari 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_februari set  3 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "04 Februari 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_februari set  4 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "05 Februari 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_februari set  5 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "06 Februari 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_februari set  6 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "07 Februari 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_februari set  7 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "08 Februari 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_februari set  8 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "09 Februari 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_februari set  9 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "10 Februari 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_februari set  10 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "11 Februari 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_februari set  11 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "12 Februari 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_februari set  12 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "13 Februari 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_februari set  13 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "14 Februari 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_februari set  14 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "15 Februari 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_februari set  15 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "16 Februari 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_februari set  16 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "17 Februari 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_februari set  17 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "18 Februari 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_februari set  18 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "19 Februari 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_februari set  19 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "20 Februari 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_februari set  20 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "21 Februari 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_februari set  21 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "22 Februari 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_februari set  22 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "23 Februari 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_februari set  23 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "24 Februari 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_februari set  24 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "25 Februari 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_februari set  25 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "26 Februari 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_februari set  26 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "27 Februari 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_februari set  27 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "28 Februari 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_februari set  28 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()


        ElseIf DateTimePicker1.Text = "01 Maret 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_maret set  1 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "02 Maret 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_maret set  2 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "03 Maret 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_maret set  3 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "04 Maret 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_maret set  4 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "05 Maret 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_maret set  5 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "06 Maret 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_maret set  6 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "07 Maret 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_maret set  7 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "08 Maret 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_maret set  8 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "09 Maret 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_maret set  9 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "10 Maret 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_maret set  10 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "11 Maret 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_maret set  11 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "12 Maret 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_maret set  12 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "13 Maret 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_maret set  13 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "14 Maret 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_maret set  14 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "15 Maret 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_maret set  15 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "16 Maret 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_maret set  16 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "17 Maret 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_maret set  17 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "18 Maret 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_maret set  18 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "19 Maret 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_maret set  19 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "20 Maret 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_maret set  20 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "21 Maret 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_maret set  21 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "22 Maret 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_maret set  22 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "23 Maret 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_maret set  23 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "24 Maret 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_maret set  24 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "25 Maret 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_maret set  25 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "26 Maret 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_maret set  26 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "27 Maret 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_maret set  27 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "28 Maret 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_maret set  28 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "29 Maret 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_maret set  29 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "30 Maret 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_maret set  30 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "31 Maret 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_maret set  31 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()


        ElseIf DateTimePicker1.Text = "01 April 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_april set  1 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "02 April 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_april set  2 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "03 April 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_april set  3 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "04 April 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_april set  4 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "05 April 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_april set  5 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "06 April 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_april set  6 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "07 April 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_april set  7 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "08 April 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_april set  8 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "09 April 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_april set  9 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "10 April 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_april set  10 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "11 April 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_april set  11 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "12 April 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_april set  12 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "13 April 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_april set  13 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "14 April 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_april set  14 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "15 April 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_april set  15 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "16 April 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_april set  16 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "17 April 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_april set  17 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "18 April 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_april set  18 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "19 April 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_april set  19 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "20 April 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_april set  20 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "21 April 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_april set  21 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "22 April 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_april set  22 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "23 April 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_april set  23 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "24 April 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_april set  24 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "25 April 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_april set  25 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "26 April 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_april set  26 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "27 April 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_april set  27 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "28 April 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_april set  28 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "29 April 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_april set  29 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "30 April 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_april set  30 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()


        ElseIf DateTimePicker1.Text = "01 Mei 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_mei set  1 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "02 Mei 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_mei set  2 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "03 Mei 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_mei set  3 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "04 Mei 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_mei set  4 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "05 Mei 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_mei set  5 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "06 Mei 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_mei set  6 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "07 Mei 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_mei set  7 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "08 Mei 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_mei set  8 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "09 Mei 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_mei set  9 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "10 Mei 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_mei set  10 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "11 Mei 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_mei set  11 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "12 Mei 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_mei set  12 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "13 Mei 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_mei set  13 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "14 Mei 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_mei set  14 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "15 Mei 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_mei set  15 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "16 Mei 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_mei set  16 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "17 Mei 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_mei set  17 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "18 Mei 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_mei set  18 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "19 Mei 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_mei set  19 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "20 Mei 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_mei set  20 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "21 Mei 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_mei set  21 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "22 Mei 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_mei set  22 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "23 Mei 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_mei set  23 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "24 Mei 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_mei set  24 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "25 Mei 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_mei set  25 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "26 Mei 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_mei set  26 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "27 Mei 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_mei set  27 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "28 Mei 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_mei set  28 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "29 Mei 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_mei set  29 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "30 Mei 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_mei set  30 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "31 Mei 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_mei set  31 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()


        ElseIf DateTimePicker1.Text = "01 Juni 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_juni set  1 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "02 Juni 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_juni set  2 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "03 Juni 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_juni set  3 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "04 Juni 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_juni set  4 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "05 Juni 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_juni set  5 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "06 Juni 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_juni set  6 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "07 Juni 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_juni set  7 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "08 Juni 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_juni set  8 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "09 Juni 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_juni set  9 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "10 Juni 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_juni set  10 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "11 Juni 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_juni set  11 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "12 Juni 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_juni set  12 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "13 Juni 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_juni set  13 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "14 Juni 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_juni set  14 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "15 Juni 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_juni set  15 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "16 Juni 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_juni set  16 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "17 Juni 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_juni set  17 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "18 Juni 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_juni set  18 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "19 Juni 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_juni set  19 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "20 Juni 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_juni set  20 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "21 Juni 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_juni set  21 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "22 Juni 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_juni set  22 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "23 Juni 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_juni set  23 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "24 Juni 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_juni set  24 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "25 Juni 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_juni set  25 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "26 Juni 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_juni set  26 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "27 Juni 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_juni set  27 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "28 Juni 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_juni set  28 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "29 Juni 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_juni set  29 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "30 Juni 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_juni set  30 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "31 Juni 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_juni set  31 ='" & Button10.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        End If
    End Sub

    Private Sub Button11_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button11.Click
        If RadioButton1.Checked = True Then
            pilihan = RadioButton1.Text
        ElseIf RadioButton1.Checked = True Then
            pilihan = RadioButton2.Text
        ElseIf RadioButton3.Checked = True Then
            pilihan = RadioButton3.Text
        ElseIf RadioButton4.Checked = True Then
            pilihan = RadioButton4.Text

        End If
        If DateTimePicker1.Text = "01 Juli 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_Juli set  1 ='" & Button11.Text & "' "
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "02 Juli 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_juli set  2 ='" & Button11.Text & "' "
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "03 Juli 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_juli set  3 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "04 Juli 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_juli set  4 ='" & Button11.Text & "' "
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "05 Juli 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_juli set  5 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "06 Juli 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_juli set  6 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "07 Juli 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_juli set  8 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "09 Juli 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_juli set  9 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "10 Juli 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_juli set  10 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "11 Juli 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_juli set  11 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "12 Juli 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_juli set  12 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "13 Juli 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_juli set  13 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "14 Juli 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_juli set  14 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "15 Juli 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_juli set  15 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "16 Juli 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_juli set  16 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "17 Juli 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_juli set  17 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "18 Juli 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_juli set  18 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "19 Juli 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_juli set  19 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "20 Juli 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_juli set  20 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "21 Juli 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_juli set  21 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "22 Juli 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_juli set  22 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "23 Juli 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_juli set  23 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "24 Juli 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_juli set  24 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "25 Juli 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_juli set  25 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "26 Juli 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_juli set 26 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "27 Juli 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_juli set  27 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "28 Juli 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_juli set  28 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "29 Juli 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_juli set  29 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "30 Juli 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_juli set  30 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "31 Juli 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_juli set  31 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()



        ElseIf DateTimePicker1.Text = "01 Agustus 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_agustus set  1 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "02 Agustus 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_agustus set  2 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "03 Agustus 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_agustus set  3 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "04 Agustus 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_agustus set  4 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "05 Agustus 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_agustus set  5 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "06 Agustus 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_agustus set  6 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "07 Agustus 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_agustus set  7 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "08 Agustus 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_agustus set  8 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "09 Agustus 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_agustus set  9 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "10 Agustus 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_agustus set  10 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "11 Agustus 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_agustus set  11 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "12 Agustus 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_agustus set  12 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "13 Agustus 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_agustus set  13 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "14 Agustus 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_agustus set  14 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "15 Agustus 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_agustus set  15 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "16 Agustus 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_agustus set  16 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "17 Agustus 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_agustus set  17 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "18 Agustus 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_agustus set  18 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "19 Agustus 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_agustus set  19 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "20 Agustus 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_agustus set  20 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "21 Agustus 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_agustus set  21 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "22 Agustus 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_agustus set  22 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "23 Agustus 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_agustus set  23 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "24 Agustus 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_agustus set  24 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "25 Agustus 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_agustus set  25 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "26 Agustus 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_agustus set  26 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "27 Agustus 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_agustus set  27 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "28 Agustus 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_agustus set  28 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "29 Agustus 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_agustus set  29 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "30 Agustus 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_agustus set  30 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "31 Agustus 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_agustus set  31 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()


        ElseIf DateTimePicker1.Text = "01 September 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_september set  1 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "02 September 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_september set  2 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "03 September 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_september set  3 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "04 September 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_september set  4 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "05 September 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_september set  5 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "06 September 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_september set  6 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "07 September 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_september set  7 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "08 September 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_september set  8 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "09 September 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_september set  9 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "10 September 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_september set  10 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "11 September 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_september set  11 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "12 September 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_september set  12 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "13 September 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_september set  13 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "14 September 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_september set  14 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "15 September 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_september set  15 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "16 September 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_september set  16 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "17 September 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_september set  17 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "18 September 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_september set  18 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "19 September 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_september set  19 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "20 September 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_september set  20 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "21 September 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_september set  21 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "22 September 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_september set  22 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "23 September 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_september set  23 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "24 September 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_september set  24 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "25 September 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_september set  25 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "26 September 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_september set  26 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "27 September 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_september set  27 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "28 September 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_september set  28 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "29 September 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_september set  29 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "30 September 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_september set  30 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()



        ElseIf DateTimePicker1.Text = "01 Oktober 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_oktober set  1 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "02 Oktober 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_oktober set  2 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "03 Oktober 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_oktober set  3 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "04 Oktober 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_oktober set  4 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "05 Oktober 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_oktober set  5 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "06 Oktober 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_oktober set  6 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "07 Oktober 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_oktober set  7 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "08 Oktober 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_oktober set  8 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "09 Oktober 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_oktober set  9 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "10 Oktober 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_oktober set  10 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "11 Oktober 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_oktober set  11 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "12 Oktober 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_oktober set  12 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "13 Oktober 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_oktober set  13 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "14 Oktober 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_oktober set  14 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "15 Oktober 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_oktober set  15 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "16 Oktober 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_oktober set  16 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "17 Oktober 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_oktober set  17 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "18 Oktober 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_oktober set  18 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "19 Oktober 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_oktober set  19 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "20 Oktober 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_oktober set  20 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "21 Oktober 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_oktober set  21 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "22 Oktober 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_oktober set  2 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "23 Oktober 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_oktober set  23 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "24 Oktober 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_oktober set  24 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "25 Oktober 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_oktober set  25 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "26 Oktober 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_oktober set  26 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "27 Oktober 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_oktober set  27 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "28 Oktober 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_oktober set  28 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "29 Oktober 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_oktober set  29 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "30 Oktober 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_oktober set  30 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "31 Oktober 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_oktober set  31 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()


        ElseIf DateTimePicker1.Text = "01 November 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_november set  1 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "02 November 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_november set  2 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "03 November 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_november set  3 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "04 November 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_november set  4 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "05 November 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_november set  5 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "06 November 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_november set  6 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "07 November 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_november set  7 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "08 November 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_november set  8 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "09 November 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_november set  9 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "10 November 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_november set  10 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "11 November 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_november set  11 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "12 November 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_november set  12 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "13 November 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_november set  13 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "14 November 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_november set  14 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "15 November 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_november set  15 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "16 November 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_november set  16 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "17 November 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_november set  17 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "18 November 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_november set  18 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "19 November 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_november set  19 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "20 November 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_november set  20 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "21 November 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_november set  21 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "22 November 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_november set  22 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "23 November 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_november set  23 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "24 November 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_november set  24 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "25 November 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_november set  25 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "26 November 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_november set  26 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "27 November 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_november set  27 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "28 November 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_november set  28 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "29 November 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_november set  29 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "30 November 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_november set  30 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()


        ElseIf DateTimePicker1.Text = "01 Desember 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_desember set  1 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "02 Desember 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_desember set  2 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "03 Desember 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_desember set  3 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "04 Desember 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_desember set  4 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "05 Desember 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_desember set  5 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "06 Desember 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_desember set  6 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "07 Desember 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_desember set  7 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "08 Desember 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_desember set  8 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "09 Desember 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_desember set  9 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "10 Desember 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_desember set  10 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "11 Desember 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_desember set  11 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "12 Desember 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_desember set  12 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "13 Desember 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_desember set  13 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "14 Desember 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_desember set  14 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "15 Desember 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_desember set  15 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "16 Desember 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_desember set  16 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "17 Desember 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_desember set  17 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "18 Desember 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_desember set  18 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "19 Desember 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_desember set  19 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "20 Desember 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_desember set  20 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "21 Desember 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_desember set  21 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "22 Desember 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_desember set  22 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "23 Desember 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_desember set  23 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "24 Desember 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_desember set  24 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "25 Desember 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_desember set  25 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "26 Desember 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_desember set  26 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "27 Desember 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_desember set  27 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "28 Desember 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_desember set  28 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "29 Desember 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_desember set  29 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "30 Desember 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_desember set  30 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "31 Desember 2016" Then
            Call koneksi()
            Dim edit As String = "update rplb_desember set  31 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()


        ElseIf DateTimePicker1.Text = "01 Januari 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_januari set  1 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "02 Januari 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_januari set  2 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "03 Januari 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_januari set  3 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "04 Januari 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_januari set  4 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "05 Januari 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_januari set  5 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "06 Januari 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_januari set  6 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "07 Januari 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_januari set  7 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "08 Januari 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_januari set  8 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "09 Januari 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_januari set  9 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "10 Januari 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_januari set  10 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "11 Januari 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_januari set  11 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "12 Januari 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_januari set  12 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "13 Januari 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_januari set  13 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "14 Januari 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_januari set  14 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "15 Januari 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_januari set  15 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "16 Januari 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_januari set  16 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "17 Januari 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_januari set  17 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "18 Januari 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_januari set  18 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "19 Januari 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_januari set  19 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "20 Januari 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_januari set  20 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "21 Januari 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_januari set  21 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "22 Januari 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_januari set  22 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "23 Januari 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_januari set  23 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "24 Januari 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_januari set  24 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "25 Januari 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_januari set  25 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "26 Januari 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_januari set  26 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "27 Januari 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_januari set  27 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "28 Januari 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_januari set  28 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "29 Januari 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_januari set  29 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "30 Januari 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_januari set  30 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "31 Januari 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_januari set  31 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()


        ElseIf DateTimePicker1.Text = "01 Februari 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_februari set  1 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "02 Februari 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_februari set  2 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "03 Februari 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_februari set  3 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "04 Februari 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_februari set  4 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "05 Februari 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_februari set  5 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "06 Februari 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_februari set  6 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "07 Februari 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_februari set  7 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "08 Februari 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_februari set  8 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "09 Februari 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_februari set  9 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "10 Februari 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_februari set  10 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "11 Februari 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_februari set  11 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "12 Februari 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_februari set  12 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "13 Februari 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_februari set  13 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "14 Februari 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_februari set  14 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "15 Februari 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_februari set  15 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "16 Februari 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_februari set  16 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "17 Februari 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_februari set  17 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "18 Februari 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_februari set  18 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "19 Februari 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_februari set  19 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "20 Februari 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_februari set  20 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "21 Februari 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_februari set  21 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "22 Februari 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_februari set  22 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "23 Februari 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_februari set  23 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "24 Februari 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_februari set  24 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "25 Februari 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_februari set  25 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "26 Februari 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_februari set  26 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "27 Februari 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_februari set  27 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "28 Februari 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_februari set  28 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()


        ElseIf DateTimePicker1.Text = "01 Maret 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_maret set  1 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "02 Maret 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_maret set  2 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "03 Maret 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_maret set  3 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "04 Maret 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_maret set  4 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "05 Maret 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_maret set  5 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "06 Maret 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_maret set  6 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "07 Maret 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_maret set  7 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "08 Maret 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_maret set  8 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "09 Maret 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_maret set  9 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "10 Maret 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_maret set  10 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "11 Maret 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_maret set  11 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "12 Maret 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_maret set  12 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "13 Maret 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_maret set  13 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "14 Maret 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_maret set  14 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "15 Maret 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_maret set  15 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "16 Maret 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_maret set  16 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "17 Maret 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_maret set  17 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "18 Maret 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_maret set  18 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "19 Maret 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_maret set  19 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "20 Maret 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_maret set  20 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "21 Maret 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_maret set  21 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "22 Maret 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_maret set  22 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "23 Maret 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_maret set  23 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "24 Maret 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_maret set  24 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "25 Maret 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_maret set  25 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "26 Maret 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_maret set  26 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "27 Maret 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_maret set  27 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "28 Maret 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_maret set  28 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "29 Maret 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_maret set  29 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "30 Maret 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_maret set  30 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "31 Maret 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_maret set  31 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()


        ElseIf DateTimePicker1.Text = "01 April 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_april set  1 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "02 April 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_april set  2 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "03 April 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_april set  3 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "04 April 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_april set  4 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "05 April 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_april set  5 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "06 April 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_april set  6 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "07 April 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_april set  7 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "08 April 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_april set  8 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "09 April 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_april set  9 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "10 April 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_april set  10 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "11 April 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_april set  11 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "12 April 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_april set  12 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "13 April 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_april set  13 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "14 April 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_april set  14 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "15 April 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_april set  15 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "16 April 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_april set  16 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "17 April 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_april set  17 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "18 April 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_april set  18 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "19 April 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_april set  19 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "20 April 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_april set  20 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "21 April 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_april set  21 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "22 April 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_april set  22 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "23 April 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_april set  23 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "24 April 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_april set  24 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "25 April 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_april set  25 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "26 April 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_april set  26 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "27 April 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_april set  27 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "28 April 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_april set  28 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "29 April 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_april set  29 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "30 April 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_april set  30 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()


        ElseIf DateTimePicker1.Text = "01 Mei 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_mei set  1 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "02 Mei 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_mei set  2 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "03 Mei 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_mei set  3 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "04 Mei 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_mei set  4 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "05 Mei 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_mei set  5 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "06 Mei 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_mei set  6 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "07 Mei 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_mei set  7 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "08 Mei 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_mei set  8 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "09 Mei 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_mei set  9 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "10 Mei 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_mei set  10 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "11 Mei 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_mei set  11 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "12 Mei 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_mei set  12 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "13 Mei 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_mei set  13 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "14 Mei 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_mei set  14 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "15 Mei 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_mei set  15 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "16 Mei 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_mei set  16 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "17 Mei 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_mei set  17 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "18 Mei 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_mei set  18 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "19 Mei 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_mei set  19 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "20 Mei 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_mei set  20 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "21 Mei 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_mei set  21 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "22 Mei 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_mei set  22 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "23 Mei 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_mei set  23 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "24 Mei 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_mei set  24 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "25 Mei 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_mei set  25 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "26 Mei 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_mei set  26 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "27 Mei 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_mei set  27 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "28 Mei 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_mei set  28 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "29 Mei 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_mei set  29 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "30 Mei 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_mei set  30 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "31 Mei 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_mei set  31 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()


        ElseIf DateTimePicker1.Text = "01 Juni 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_juni set  1 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "02 Juni 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_juni set  2 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "03 Juni 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_juni set  3 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "04 Juni 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_juni set  4 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "05 Juni 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_juni set  5 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "06 Juni 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_juni set  6 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "07 Juni 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_juni set  7 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "08 Juni 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_juni set  8 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "09 Juni 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_juni set  9 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "10 Juni 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_juni set  10 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "11 Juni 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_juni set  11 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "12 Juni 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_juni set  12 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "13 Juni 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_juni set  13 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "14 Juni 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_juni set  14 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "15 Juni 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_juni set  15 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "16 Juni 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_juni set  16 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "17 Juni 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_juni set  17 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "18 Juni 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_juni set  18 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "19 Juni 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_juni set  19 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "20 Juni 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_juni set  20 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "21 Juni 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_juni set  21 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "22 Juni 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_juni set  22 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "23 Juni 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_juni set  23 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "24 Juni 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_juni set  24 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "25 Juni 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_juni set  25 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "26 Juni 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_juni set  26 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "27 Juni 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_juni set  27 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "28 Juni 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_juni set  28 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "29 Juni 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_juni set  29 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "30 Juni 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_juni set  30 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        ElseIf DateTimePicker1.Text = "31 Juni 2017" Then
            Call koneksi()
            Dim edit As String = "update rplb_juni set  31 ='" & Button11.Text & "'"
            cmd = New OleDbCommand(edit, conn)
            cmd.ExecuteNonQuery()
            Call tampilgrid1()
        End If
    End Sub

    Private Sub Button12_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button12.Click
        Call kosong2()
    End Sub

    Private Sub Button13_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        If SaveFileDialog1.ShowDialog = Windows.Forms.DialogResult.OK Then
            TextBox6.Text = SaveFileDialog1.FileName
        End If
    End Sub

    Private Sub Button14_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button14.Click
        If SaveFileDialog1.ShowDialog = Windows.Forms.DialogResult.OK Then
            TextBox6.Text = SaveFileDialog1.FileName
        End If
        Dim excelapp As vbexcel.Application
        Dim excelworkbook As vbexcel.Workbook
        Dim excelworksheet As vbexcel.Worksheet
        Dim msvalue As Object = System.Reflection.Missing.Value

        excelapp = New vbexcel.Application
        excelworkbook = excelapp.Workbooks.Add(msvalue)
        excelworksheet = excelworkbook.Sheets("sheet1")

        'Perulangan untuk memindahkan data dari datagridview ke worksheet excel
        For i As Integer = 0 To dgv.RowCount - 2
            For j As Integer = 0 To dgv.ColumnCount - 1
                If i = 0 Then
                    excelworksheet.Cells(i + 4, j + 2) = dgv.Columns(j).HeaderText.ToString
                Else
                    excelworksheet.Cells(i + 4, j + 2) = dgv(j, i).Value.ToString
                End If
            Next
        Next

        ' Mengatur ukuran baris dan column
        excelworksheet.UsedRange.EntireRow.AutoFit()
        excelworksheet.UsedRange.EntireColumn.AutoFit()

        ' Mengatur border table
        excelworksheet.UsedRange.Borders.LineStyle = vbexcel.XlLineStyle.xlContinuous
        excelworksheet.UsedRange.Borders.Color = Color.Black
        excelworksheet.UsedRange.Borders.Weight = vbexcel.XlBorderWeight.xlThin

        ' Membuat judul
        excelworksheet.Cells(1, 2).Font.Bold = True
        excelworksheet.Cells(1, 2).Font.Size = 20

        excelworksheet.Cells(1, 2) = "Data Siswa"


        ' Menyimpan file excel
        excelworksheet.SaveAs(TextBox6.Text)
        excelworkbook.Close()
        excelapp.Quit()

        MessageBox.Show("Data berhasil diexport", "Informasi", MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub

    Private Sub Button13_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button13.Click
        If SaveFileDialog1.ShowDialog = Windows.Forms.DialogResult.OK Then
            TextBox6.Text = SaveFileDialog1.FileName
        End If
        Dim excelapp As vbexcel.Application
        Dim excelworkbook As vbexcel.Workbook
        Dim excelworksheet As vbexcel.Worksheet
        Dim msvalue As Object = System.Reflection.Missing.Value

        excelapp = New vbexcel.Application
        excelworkbook = excelapp.Workbooks.Add(msvalue)
        excelworksheet = excelworkbook.Sheets("sheet1")

        'Perulangan untuk memindahkan data dari datagridview ke worksheet excel
        For i As Integer = 0 To dgv.RowCount - 2
            For j As Integer = 0 To dgv2.ColumnCount - 1
                If i = 0 Then
                    excelworksheet.Cells(i + 4, j + 2) = dgv2.Columns(j).HeaderText.ToString
                Else
                    excelworksheet.Cells(i + 4, j + 2) = dgv2(j, i).Value.ToString
                End If
            Next
        Next

        ' Mengatur ukuran baris dan column
        excelworksheet.UsedRange.EntireRow.AutoFit()
        excelworksheet.UsedRange.EntireColumn.AutoFit()

        ' Mengatur border table
        excelworksheet.UsedRange.Borders.LineStyle = vbexcel.XlLineStyle.xlContinuous
        excelworksheet.UsedRange.Borders.Color = Color.Black
        excelworksheet.UsedRange.Borders.Weight = vbexcel.XlBorderWeight.xlThin

        ' Membuat judul
        excelworksheet.Cells(1, 2).Font.Bold = True
        excelworksheet.Cells(1, 2).Font.Size = 20

        excelworksheet.Cells(1, 2) = "Data Siswa"


        ' Menyimpan file excel
        excelworksheet.SaveAs(TextBox6.Text)
        excelworkbook.Close()
        excelapp.Quit()

        MessageBox.Show("Data berhasil diexport", "Informasi", MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub

    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
        edit2.Show()
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click

            If DateTimePicker1.Text = "01 Juli 2016" Or DateTimePicker1.Text = "02 Juli 2016" Or DateTimePicker1.Text = "03 Juli 2016" Or DateTimePicker1.Text = "04 Juli 2016" Or DateTimePicker1.Text = "05 Juli 2016" Or DateTimePicker1.Text = "06 Juli 2016" Or DateTimePicker1.Text = "07 Juli 2016" Or DateTimePicker1.Text = "08 Juli 2016" Or DateTimePicker1.Text = "09 Juli 2016" Or DateTimePicker1.Text = "10 Juli 2016" Or DateTimePicker1.Text = "11 Juli 2016" Or DateTimePicker1.Text = "12 Juli 2016" Or DateTimePicker1.Text = "13 Juli 2016" Or DateTimePicker1.Text = "14 Juli 2016" Or DateTimePicker1.Text = "15 Juli 2016" Or DateTimePicker1.Text = "16 Juli 2016" Or DateTimePicker1.Text = "17 Juli 2016" Or DateTimePicker1.Text = "18 Juli 2016" Or DateTimePicker1.Text = "19 Juli 2016" Or DateTimePicker1.Text = "20 Juli 2016" Or DateTimePicker1.Text = "21 Juli 2016" Or DateTimePicker1.Text = "22 Juli 2016" Or DateTimePicker1.Text = "23 Juli 2016" Or DateTimePicker1.Text = "24 Juli 2016" Or DateTimePicker1.Text = "25 Juli 2016" Or DateTimePicker1.Text = "26 Juli 2016" Or DateTimePicker1.Text = "27 Juli 2016" Or DateTimePicker1.Text = "28 Juli 2016" Or DateTimePicker1.Text = "29 Juli 2016" Or DateTimePicker1.Text = "30 Juli 2016" Or DateTimePicker1.Text = "31 Juli 2016" Then

                If TextBox3.Text = "" Then
                    MsgBox("Data ada yang kosong")
                    Exit Sub
                End If
                Call carikode2()
                If Not dr.HasRows Then
                    MsgBox("NIS yang anda masukkan tidak terdaftar")
                    Exit Sub
                End If
                If MessageBox.Show("Yakin ingin hapus data ini...??", "", MessageBoxButtons.YesNo) = Windows.Forms.DialogResult.Yes Then
                    Call koneksi()
                    Dim hapus As String = "delete from rplb_juli where NIS= '" & TextBox3.Text & "'"
                    cmd = New OleDbCommand(hapus, conn)
                    cmd.ExecuteNonQuery()

                    Call tampilgrid1()
            End If
            ElseIf DateTimePicker1.Text = "01 Agustus 2016" Or DateTimePicker1.Text = "02 Agustus 2016" Or DateTimePicker1.Text = "03 Agustus 2016" Or DateTimePicker1.Text = "04 Agustus 2016" Or DateTimePicker1.Text = "05 Agustus 2016" Or DateTimePicker1.Text = "06 Agustus 2016" Or DateTimePicker1.Text = "07 Agustus 2016" Or DateTimePicker1.Text = "08 Agustus 2016" Or DateTimePicker1.Text = "09 Agustus 2016" Or DateTimePicker1.Text = "10 Agustus 2016" Or DateTimePicker1.Text = "11 Agustus 2016" Or DateTimePicker1.Text = "12 Agustus 2016" Or DateTimePicker1.Text = "13 Agustus 2016" Or DateTimePicker1.Text = "14 Agustus 2016" Or DateTimePicker1.Text = "15 Agustus 2016" Or DateTimePicker1.Text = "16 Agustus 2016" Or DateTimePicker1.Text = "17 Agustus 2016" Or DateTimePicker1.Text = "18 Agustus 2016" Or DateTimePicker1.Text = "19 Agustus 2016" Or DateTimePicker1.Text = "20 Agustus 2016" Or DateTimePicker1.Text = "21 Agustus 2016" Or DateTimePicker1.Text = "22 Agustus 2016" Or DateTimePicker1.Text = "23 Agustus 2016" Or DateTimePicker1.Text = "24 Agustus 2016" Or DateTimePicker1.Text = "25 Agustus 2016" Or DateTimePicker1.Text = "26 Agustus 2016" Or DateTimePicker1.Text = "27 Agustus 2016" Or DateTimePicker1.Text = "28 Agustus 2016" Or DateTimePicker1.Text = "29 Agustus 2016" Or DateTimePicker1.Text = "30 Agustus 2016" Or DateTimePicker1.Text = "31 Agustus 2016" Then
                If TextBox3.Text = "" Then
                    MsgBox("Data ada yang kosong")
                    Exit Sub
                End If
                Call carikode2()
                If Not dr.HasRows Then
                    MsgBox("NIS yang anda masukkan tidak terdaftar")
                    Exit Sub
                End If
                If MessageBox.Show("Yakin ingin hapus data ini...??", "", MessageBoxButtons.YesNo) = Windows.Forms.DialogResult.Yes Then
                    Call koneksi()
                    Dim hapus As String = "delete from rplb_agustus where NIS= '" & TextBox3.Text & "'"
                    cmd = New OleDbCommand(hapus, conn)
                    cmd.ExecuteNonQuery()

                    Call tampilgrid1()
                End If
            ElseIf DateTimePicker1.Text = "01 September 2016" Or DateTimePicker1.Text = "02 September 2016" Or DateTimePicker1.Text = "03 September 2016" Or DateTimePicker1.Text = "04 September 2016" Or DateTimePicker1.Text = "05 September 2016" Or DateTimePicker1.Text = "06 September 2016" Or DateTimePicker1.Text = "07 September 2016" Or DateTimePicker1.Text = "08 September 2016" Or DateTimePicker1.Text = "09 September 2016" Or DateTimePicker1.Text = "10 September 2016" Or DateTimePicker1.Text = "11 September 2016" Or DateTimePicker1.Text = "12 September 2016" Or DateTimePicker1.Text = "13 September 2016" Or DateTimePicker1.Text = "14 September 2016" Or DateTimePicker1.Text = "15 September 2016" Or DateTimePicker1.Text = "16 September 2016" Or DateTimePicker1.Text = "17 September 2016" Or DateTimePicker1.Text = "18 September 2016" Or DateTimePicker1.Text = "19 September 2016" Or DateTimePicker1.Text = "20 September 2016" Or DateTimePicker1.Text = "21 September 2016" Or DateTimePicker1.Text = "22 September 2016" Or DateTimePicker1.Text = "23 September 2016" Or DateTimePicker1.Text = "24 September 2016" Or DateTimePicker1.Text = "25 September 2016" Or DateTimePicker1.Text = "26 September 2016" Or DateTimePicker1.Text = "27 September 2016" Or DateTimePicker1.Text = "28 September 2016" Or DateTimePicker1.Text = "29 September 2016" Or DateTimePicker1.Text = "30 September 2016" Then
                If TextBox3.Text = "" Then
                    MsgBox("Data ada yang kosong")
                    Exit Sub
                End If
                Call carikode2()
                If Not dr.HasRows Then
                    MsgBox("NIS yang anda masukkan tidak terdaftar")
                    Exit Sub
                End If
                If MessageBox.Show("Yakin ingin hapus data ini...??", "", MessageBoxButtons.YesNo) = Windows.Forms.DialogResult.Yes Then
                    Call koneksi()
                    Dim hapus As String = "delete from rplb_september where NIS= '" & TextBox3.Text & "'"
                    cmd = New OleDbCommand(hapus, conn)
                    cmd.ExecuteNonQuery()

                    Call tampilgrid1()
                End If
            ElseIf DateTimePicker1.Text = "01 Oktober 2016" Or DateTimePicker1.Text = "02 Oktober 2016" Or DateTimePicker1.Text = "03 Oktober 2016" Or DateTimePicker1.Text = "04 Oktober 2016" Or DateTimePicker1.Text = "05 Oktober 2016" Or DateTimePicker1.Text = "06 Oktober 2016" Or DateTimePicker1.Text = "07 Oktober 2016" Or DateTimePicker1.Text = "08 Oktober 2016" Or DateTimePicker1.Text = "09 Oktober 2016" Or DateTimePicker1.Text = "10 Oktober 2016" Or DateTimePicker1.Text = "11 Oktober 2016" Or DateTimePicker1.Text = "12 Oktober 2016" Or DateTimePicker1.Text = "13 Oktober 2016" Or DateTimePicker1.Text = "14 Oktober 2016" Or DateTimePicker1.Text = "15 Oktober 2016" Or DateTimePicker1.Text = "16 Oktober 2016" Or DateTimePicker1.Text = "17 Oktober 2016" Or DateTimePicker1.Text = "18 Oktober 2016" Or DateTimePicker1.Text = "19 Oktober 2016" Or DateTimePicker1.Text = "20 Oktober 2016" Or DateTimePicker1.Text = "21 Oktober 2016" Or DateTimePicker1.Text = "22 Oktober 2016" Or DateTimePicker1.Text = "23 Oktober 2016" Or DateTimePicker1.Text = "24 Oktober 2016" Or DateTimePicker1.Text = "25 Oktober 2016" Or DateTimePicker1.Text = "26 Oktober 2016" Or DateTimePicker1.Text = "27 Oktober 2016" Or DateTimePicker1.Text = "28 Oktober 2016" Or DateTimePicker1.Text = "29 Oktober 2016" Or DateTimePicker1.Text = "30 Oktober 2016" Or DateTimePicker1.Text = "31 Oktober 2016" Then
                If TextBox3.Text = "" Then
                    MsgBox("Data ada yang kosong")
                    Exit Sub
                End If
                Call carikode2()
                If Not dr.HasRows Then
                    MsgBox("NIS yang anda masukkan tidak terdaftar")
                    Exit Sub
                End If
                If MessageBox.Show("Yakin ingin hapus data ini...??", "", MessageBoxButtons.YesNo) = Windows.Forms.DialogResult.Yes Then
                    Call koneksi()
                    Dim hapus As String = "delete from rplb_oktober where NIS= '" & TextBox3.Text & "'"
                    cmd = New OleDbCommand(hapus, conn)
                    cmd.ExecuteNonQuery()

                    Call tampilgrid1()
                End If
            ElseIf DateTimePicker1.Text = "01 November 2016" Or DateTimePicker1.Text = "02 November 2016" Or DateTimePicker1.Text = "03 November 2016" Or DateTimePicker1.Text = "04 November 2016" Or DateTimePicker1.Text = "05 November 2016" Or DateTimePicker1.Text = "06 November 2016" Or DateTimePicker1.Text = "07 November 2016" Or DateTimePicker1.Text = "08 November 2016" Or DateTimePicker1.Text = "09 November 2016" Or DateTimePicker1.Text = "10 November 2016" Or DateTimePicker1.Text = "11 November 2016" Or DateTimePicker1.Text = "12 November 2016" Or DateTimePicker1.Text = "13 November 2016" Or DateTimePicker1.Text = "14 November 2016" Or DateTimePicker1.Text = "15 November 2016" Or DateTimePicker1.Text = "16 November 2016" Or DateTimePicker1.Text = "17 November 2016" Or DateTimePicker1.Text = "18 November 2016" Or DateTimePicker1.Text = "19 November 2016" Or DateTimePicker1.Text = "20 November 2016" Or DateTimePicker1.Text = "21 November 2016" Or DateTimePicker1.Text = "22 November 2016" Or DateTimePicker1.Text = "23 November 2016" Or DateTimePicker1.Text = "24 November 2016" Or DateTimePicker1.Text = "25 November 2016" Or DateTimePicker1.Text = "26 November 2016" Or DateTimePicker1.Text = "27 November 2016" Or DateTimePicker1.Text = "28 November 2016" Or DateTimePicker1.Text = "29 November 2016" Or DateTimePicker1.Text = "30 November 2016" Then
                If TextBox3.Text = "" Then
                    MsgBox("Data ada yang kosong")
                    Exit Sub
                End If
                Call carikode2()
                If Not dr.HasRows Then
                    MsgBox("NIS yang anda masukkan tidak terdaftar")
                    Exit Sub
                End If
                If MessageBox.Show("Yakin ingin hapus data ini...??", "", MessageBoxButtons.YesNo) = Windows.Forms.DialogResult.Yes Then
                    Call koneksi()
                    Dim hapus As String = "delete from rplb_november where NIS= '" & TextBox3.Text & "'"
                    cmd = New OleDbCommand(hapus, conn)
                    cmd.ExecuteNonQuery()

                    Call tampilgrid1()
                End If
            ElseIf DateTimePicker1.Text = "01 Desember 2016" Or DateTimePicker1.Text = "02 Desember 2016" Or DateTimePicker1.Text = "03 Desember 2016" Or DateTimePicker1.Text = "04 Desember 2016" Or DateTimePicker1.Text = "05 Desember 2016" Or DateTimePicker1.Text = "06 Desember 2016" Or DateTimePicker1.Text = "07 Desember 2016" Or DateTimePicker1.Text = "08 Desember 2016" Or DateTimePicker1.Text = "09 Desember 2016" Or DateTimePicker1.Text = "10 Desember 2016" Or DateTimePicker1.Text = "11 Desember 2016" Or DateTimePicker1.Text = "12 Desember 2016" Or DateTimePicker1.Text = "13 Desember 2016" Or DateTimePicker1.Text = "14 Desember 2016" Or DateTimePicker1.Text = "15 Desember 2016" Or DateTimePicker1.Text = "16 Desember 2016" Or DateTimePicker1.Text = "17 Desember 2016" Or DateTimePicker1.Text = "18 Desember 2016" Or DateTimePicker1.Text = "19 Desember 2016" Or DateTimePicker1.Text = "20 Desember 2016" Or DateTimePicker1.Text = "21 Desember 2016" Or DateTimePicker1.Text = "22 Desember 2016" Or DateTimePicker1.Text = "23 Desember 2016" Or DateTimePicker1.Text = "24 Desember 2016" Or DateTimePicker1.Text = "25 Desember 2016" Or DateTimePicker1.Text = "26 Desember 2016" Or DateTimePicker1.Text = "27 Desember 2016" Or DateTimePicker1.Text = "28 Desember 2016" Or DateTimePicker1.Text = "29 Desember 2016" Or DateTimePicker1.Text = "30 Desember 2016" Or DateTimePicker1.Text = "31 Desember 2016" Then
                If TextBox3.Text = "" Then
                    MsgBox("Data ada yang kosong")
                    Exit Sub
                End If
                Call carikode2()
                If Not dr.HasRows Then
                    MsgBox("NIS yang anda masukkan tidak terdaftar")
                    Exit Sub
                End If
                If MessageBox.Show("Yakin ingin hapus data ini...??", "", MessageBoxButtons.YesNo) = Windows.Forms.DialogResult.Yes Then
                    Call koneksi()
                    Dim hapus As String = "delete from rplb_desember where NIS= '" & TextBox3.Text & "'"
                    cmd = New OleDbCommand(hapus, conn)
                    cmd.ExecuteNonQuery()

                    Call tampilgrid1()
                End If
            ElseIf DateTimePicker1.Text = "01 Januari 2017" Or DateTimePicker1.Text = "02 Januari 2017" Or DateTimePicker1.Text = "03 Januari 2017" Or DateTimePicker1.Text = "04 Januari 2017" Or DateTimePicker1.Text = "05 Januari 2017" Or DateTimePicker1.Text = "06 Januari 2017" Or DateTimePicker1.Text = "07 Januari 2017" Or DateTimePicker1.Text = "08 Januari 2017" Or DateTimePicker1.Text = "09 Januari 2017" Or DateTimePicker1.Text = "10 Januari 2017" Or DateTimePicker1.Text = "11 Januari 2017" Or DateTimePicker1.Text = "12 Januari 2017" Or DateTimePicker1.Text = "13 Januari 2017" Or DateTimePicker1.Text = "14 Januari 2017" Or DateTimePicker1.Text = "15 Januari 2017" Or DateTimePicker1.Text = "16 Januari 2017" Or DateTimePicker1.Text = "17 Januari 2017" Or DateTimePicker1.Text = "18 Januari 2017" Or DateTimePicker1.Text = "19 Januari 2017" Or DateTimePicker1.Text = "20 Januari 2017" Or DateTimePicker1.Text = "21 Januari 2017" Or DateTimePicker1.Text = "22 Januari 2017" Or DateTimePicker1.Text = "23 Januari 2017" Or DateTimePicker1.Text = "24 Januari 2017" Or DateTimePicker1.Text = "25 Januari 2017" Or DateTimePicker1.Text = "26 Januari 2017" Or DateTimePicker1.Text = "27 Januari 2017" Or DateTimePicker1.Text = "28 Januari 2017" Or DateTimePicker1.Text = "29 Januari 2017" Or DateTimePicker1.Text = "30 Januari 2017" Or DateTimePicker1.Text = "31 Januari 2017" Then
                If TextBox3.Text = "" Then
                    MsgBox("Data ada yang kosong")
                    Exit Sub
                End If
                Call carikode2()
                If Not dr.HasRows Then
                    MsgBox("NIS yang anda masukkan tidak terdaftar")
                    Exit Sub
                End If
                If MessageBox.Show("Yakin ingin hapus data ini...??", "", MessageBoxButtons.YesNo) = Windows.Forms.DialogResult.Yes Then
                    Call koneksi()
                    Dim hapus As String = "delete from rplb_januari where NIS= '" & TextBox3.Text & "'"
                    cmd = New OleDbCommand(hapus, conn)
                    cmd.ExecuteNonQuery()

                    Call tampilgrid1()
                End If
            ElseIf DateTimePicker1.Text = "01 Februari 2017" Or DateTimePicker1.Text = "02 Februari 2017" Or DateTimePicker1.Text = "03 Februari 2017" Or DateTimePicker1.Text = "04 Februari 2017" Or DateTimePicker1.Text = "05 Februari 2017" Or DateTimePicker1.Text = "06 Februari 2017" Or DateTimePicker1.Text = "07 Februari 2017" Or DateTimePicker1.Text = "08 Februari 2017" Or DateTimePicker1.Text = "09 Februari 2017" Or DateTimePicker1.Text = "10 Februari 2017" Or DateTimePicker1.Text = "11 Februari 2017" Or DateTimePicker1.Text = "12 Februari 2017" Or DateTimePicker1.Text = "13 Februari 2017" Or DateTimePicker1.Text = "14 Februari 2017" Or DateTimePicker1.Text = "15 Februari 2017" Or DateTimePicker1.Text = "16 Februari 2017" Or DateTimePicker1.Text = "17 Februari 2017" Or DateTimePicker1.Text = "18 Februari 2017" Or DateTimePicker1.Text = "19 Februari 2017" Or DateTimePicker1.Text = "20 Februari 2017" Or DateTimePicker1.Text = "21 Februari 2017" Or DateTimePicker1.Text = "22 Februari 2017" Or DateTimePicker1.Text = "23 Februari 2017" Or DateTimePicker1.Text = "24 Februari 2017" Or DateTimePicker1.Text = "25 Februari 2017" Or DateTimePicker1.Text = "26 Februari 2017" Or DateTimePicker1.Text = "27 Februari 2017" Or DateTimePicker1.Text = "28 Februari 2017" Then
                If TextBox3.Text = "" Then
                    MsgBox("Data ada yang kosong")
                    Exit Sub
                End If
                Call carikode2()
                If Not dr.HasRows Then
                    MsgBox("NIS yang anda masukkan tidak terdaftar")
                    Exit Sub
                End If
                If MessageBox.Show("Yakin ingin hapus data ini...??", "", MessageBoxButtons.YesNo) = Windows.Forms.DialogResult.Yes Then
                    Call koneksi()
                    Dim hapus As String = "delete from rplb_februari where NIS= '" & TextBox3.Text & "'"
                    cmd = New OleDbCommand(hapus, conn)
                    cmd.ExecuteNonQuery()

                    Call tampilgrid1()
                End If
            ElseIf DateTimePicker1.Text = "01 Maret 2017" Or DateTimePicker1.Text = "02 Maret 2017" Or DateTimePicker1.Text = "03 Maret 2017" Or DateTimePicker1.Text = "04 Maret 2017" Or DateTimePicker1.Text = "05 Maret 2017" Or DateTimePicker1.Text = "06 Maret 2017" Or DateTimePicker1.Text = "07 Maret 2017" Or DateTimePicker1.Text = "08 Maret 2017" Or DateTimePicker1.Text = "09 Maret 2017" Or DateTimePicker1.Text = "10 Maret 2017" Or DateTimePicker1.Text = "11 Maret 2017" Or DateTimePicker1.Text = "12 Maret 2017" Or DateTimePicker1.Text = "13 Maret 2017" Or DateTimePicker1.Text = "14 Maret 2017" Or DateTimePicker1.Text = "15 Maret 2017" Or DateTimePicker1.Text = "16 Maret 2017" Or DateTimePicker1.Text = "17 Maret 2017" Or DateTimePicker1.Text = "18 Maret 2017" Or DateTimePicker1.Text = "19 Maret 2017" Or DateTimePicker1.Text = "20 Maret 2017" Or DateTimePicker1.Text = "21 Maret 2017" Or DateTimePicker1.Text = "22 Maret 2017" Or DateTimePicker1.Text = "23 Maret 2017" Or DateTimePicker1.Text = "24 Maret 2017" Or DateTimePicker1.Text = "25 Maret 2017" Or DateTimePicker1.Text = "26 Maret 2017" Or DateTimePicker1.Text = "27 Maret 2017" Or DateTimePicker1.Text = "28 Maret 2017" Or DateTimePicker1.Text = "29 Maret 2017" Or DateTimePicker1.Text = "30 Maret 2017" Or DateTimePicker1.Text = "31 Maret 2017" Then
                If TextBox3.Text = "" Then
                    MsgBox("Data ada yang kosong")
                    Exit Sub
                End If
                Call carikode2()
                If Not dr.HasRows Then
                    MsgBox("NIS yang anda masukkan tidak terdaftar")
                    Exit Sub
                End If
                If MessageBox.Show("Yakin ingin hapus data ini...??", "", MessageBoxButtons.YesNo) = Windows.Forms.DialogResult.Yes Then
                    Call koneksi()
                    Dim hapus As String = "delete from rplb_maret where NIS= '" & TextBox3.Text & "'"
                    cmd = New OleDbCommand(hapus, conn)
                    cmd.ExecuteNonQuery()

                    Call tampilgrid1()
                End If
            ElseIf DateTimePicker1.Text = "01 April 2017" Or DateTimePicker1.Text = "02 April 2017" Or DateTimePicker1.Text = "03 April 2017" Or DateTimePicker1.Text = "04 April 2017" Or DateTimePicker1.Text = "05 April 2017" Or DateTimePicker1.Text = "06 April 2017" Or DateTimePicker1.Text = "07 April 2017" Or DateTimePicker1.Text = "08 April 2017" Or DateTimePicker1.Text = "09 April 2017" Or DateTimePicker1.Text = "10 April 2017" Or DateTimePicker1.Text = "11 April 2017" Or DateTimePicker1.Text = "12 April 2017" Or DateTimePicker1.Text = "13 April 2017" Or DateTimePicker1.Text = "14 April 2017" Or DateTimePicker1.Text = "15 April 2017" Or DateTimePicker1.Text = "16 April 2017" Or DateTimePicker1.Text = "17 April 2017" Or DateTimePicker1.Text = "18 April 2017" Or DateTimePicker1.Text = "19 April 2017" Or DateTimePicker1.Text = "20 April 2017" Or DateTimePicker1.Text = "21 April 2017" Or DateTimePicker1.Text = "22 April 2017" Or DateTimePicker1.Text = "23 April 2017" Or DateTimePicker1.Text = "24 April 2017" Or DateTimePicker1.Text = "25 April 2017" Or DateTimePicker1.Text = "26 April 2017" Or DateTimePicker1.Text = "27 April 2017" Or DateTimePicker1.Text = "28 April 2017" Or DateTimePicker1.Text = "29 April 2017" Or DateTimePicker1.Text = "30 April 2017" Then
                If TextBox3.Text = "" Then
                    MsgBox("Data ada yang kosong")
                    Exit Sub
                End If
                Call carikode2()
                If Not dr.HasRows Then
                    MsgBox("NIS yang anda masukkan tidak terdaftar")
                    Exit Sub
                End If
                If MessageBox.Show("Yakin ingin hapus data ini...??", "", MessageBoxButtons.YesNo) = Windows.Forms.DialogResult.Yes Then
                    Call koneksi()
                    Dim hapus As String = "delete from rplb_april where NIS= '" & TextBox3.Text & "'"
                    cmd = New OleDbCommand(hapus, conn)
                    cmd.ExecuteNonQuery()

                    Call tampilgrid1()
                End If
            ElseIf DateTimePicker1.Text = "01 Mei 2017" Or DateTimePicker1.Text = "02 Mei 2017" Or DateTimePicker1.Text = "03 Mei 2017" Or DateTimePicker1.Text = "04 Mei 2017" Or DateTimePicker1.Text = "05 Mei 2017" Or DateTimePicker1.Text = "06 Mei 2017" Or DateTimePicker1.Text = "07 Mei 2017" Or DateTimePicker1.Text = "08 Mei 2017" Or DateTimePicker1.Text = "09 Mei 2017" Or DateTimePicker1.Text = "10 Mei 2017" Or DateTimePicker1.Text = "11 Mei 2017" Or DateTimePicker1.Text = "12 Mei 2017" Or DateTimePicker1.Text = "13 Mei 2017" Or DateTimePicker1.Text = "14 Mei 2017" Or DateTimePicker1.Text = "15 Mei 2017" Or DateTimePicker1.Text = "16 Mei 2017" Or DateTimePicker1.Text = "17 Mei 2017" Or DateTimePicker1.Text = "18 Mei 2017" Or DateTimePicker1.Text = "19 Mei 2017" Or DateTimePicker1.Text = "20 Mei 2017" Or DateTimePicker1.Text = "21 Mei 2017" Or DateTimePicker1.Text = "22 Mei 2017" Or DateTimePicker1.Text = "23 Mei 2017" Or DateTimePicker1.Text = "24 Mei 2017" Or DateTimePicker1.Text = "25 Mei 2017" Or DateTimePicker1.Text = "26 Mei 2017" Or DateTimePicker1.Text = "27 Mei 2017" Or DateTimePicker1.Text = "28 Mei 2017" Or DateTimePicker1.Text = "29 Mei 2017" Or DateTimePicker1.Text = "30 Mei 2017" Or DateTimePicker1.Text = "31 Mei 2017" Then
                If TextBox3.Text = "" Then
                    MsgBox("Data ada yang kosong")
                    Exit Sub
                End If
                Call carikode2()
                If Not dr.HasRows Then
                    MsgBox("NIS yang anda masukkan tidak terdaftar")
                    Exit Sub
                End If
                If MessageBox.Show("Yakin ingin hapus data ini...??", "", MessageBoxButtons.YesNo) = Windows.Forms.DialogResult.Yes Then
                    Call koneksi()
                    Dim hapus As String = "delete from rplb_mei where NIS= '" & TextBox3.Text & "'"
                    cmd = New OleDbCommand(hapus, conn)
                    cmd.ExecuteNonQuery()

                    Call tampilgrid1()
                End If
            ElseIf DateTimePicker1.Text = "01 Juni 2017" Or DateTimePicker1.Text = "02 Juni 2017" Or DateTimePicker1.Text = "03 Juni 2017" Or DateTimePicker1.Text = "04 Juni 2017" Or DateTimePicker1.Text = "05 Juni 2017" Or DateTimePicker1.Text = "06 Juni 2017" Or DateTimePicker1.Text = "07 Juni 2017" Or DateTimePicker1.Text = "08 Juni 2017" Or DateTimePicker1.Text = "09 Juni 2017" Or DateTimePicker1.Text = "10 Juni 2017" Or DateTimePicker1.Text = "11 Juni 2017" Or DateTimePicker1.Text = "12 Juni 2017" Or DateTimePicker1.Text = "13 Juni 2017" Or DateTimePicker1.Text = "14 Juni 2017" Or DateTimePicker1.Text = "15 Juni 2017" Or DateTimePicker1.Text = "16 Juni 2017" Or DateTimePicker1.Text = "17 Juni 2017" Or DateTimePicker1.Text = "18 Juni 2017" Or DateTimePicker1.Text = "19 Juni 2017" Or DateTimePicker1.Text = "20 Juni 2017" Or DateTimePicker1.Text = "21 Juni 2017" Or DateTimePicker1.Text = "22 Juni 2017" Or DateTimePicker1.Text = "23 Juni 2017" Or DateTimePicker1.Text = "24 Juni 2017" Or DateTimePicker1.Text = "25 Juni 2017" Or DateTimePicker1.Text = "26 Juni 2017" Or DateTimePicker1.Text = "27 Juni 2017" Or DateTimePicker1.Text = "28 Juni 2017" Or DateTimePicker1.Text = "29 Juni 2017" Or DateTimePicker1.Text = "30 Juni 2017" Then
                If TextBox3.Text = "" Then
                    MsgBox("Data ada yang kosong")
                    Exit Sub
                End If
                Call carikode2()
                If Not dr.HasRows Then
                    MsgBox("NIS yang anda masukkan tidak terdaftar")
                    Exit Sub
                End If
                If MessageBox.Show("Yakin ingin hapus data ini...??", "", MessageBoxButtons.YesNo) = Windows.Forms.DialogResult.Yes Then
                    Call koneksi()
                    Dim hapus As String = "delete from rplb_juni where NIS= '" & TextBox3.Text & "'"
                    cmd = New OleDbCommand(hapus, conn)
                    cmd.ExecuteNonQuery()

                    Call tampilgrid1()
                End If

            End If

    End Sub

    Private Sub Label2_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label2.Click

    End Sub
End Class
