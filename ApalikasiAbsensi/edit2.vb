﻿Imports System.Data.OleDb
Public Class edit2

    Sub ketemu3()
        TextBox1.Text = dr.Item(0)
        TextBox2.Text = dr.Item(1)

    End Sub

    Sub carikode2()
        If Form1.DateTimePicker1.Text = "01 Juli 2016" Or Form1.DateTimePicker1.Text = "02 Juli 2016" Or Form1.DateTimePicker1.Text = "03 Juli 2016" Or Form1.DateTimePicker1.Text = "04 Juli 2016" Or Form1.DateTimePicker1.Text = "05 Juli 2016" Or Form1.DateTimePicker1.Text = "06 Juli 2016" Or Form1.DateTimePicker1.Text = "07 Juli 2016" Or Form1.DateTimePicker1.Text = "08 Juli 2016" Or Form1.DateTimePicker1.Text = "09 Juli 2016" Or Form1.DateTimePicker1.Text = "10 Juli 2016" Or Form1.DateTimePicker1.Text = "11 Juli 2016" Or Form1.DateTimePicker1.Text = "12 Juli 2016" Or Form1.DateTimePicker1.Text = "13 Juli 2016" Or Form1.DateTimePicker1.Text = "14 Juli 2016" Or Form1.DateTimePicker1.Text = "15 Juli 2016" Or Form1.DateTimePicker1.Text = "16 Juli 2016" Or Form1.DateTimePicker1.Text = "17 Juli 2016" Or Form1.DateTimePicker1.Text = "18 Juli 2016" Or Form1.DateTimePicker1.Text = "19 Juli 2016" Or Form1.DateTimePicker1.Text = "20 Juli 2016" Or Form1.DateTimePicker1.Text = "21 Juli 2016" Or Form1.DateTimePicker1.Text = "22 Juli 2016" Or Form1.DateTimePicker1.Text = "23 Juli 2016" Or Form1.DateTimePicker1.Text = "24 Juli 2016" Or Form1.DateTimePicker1.Text = "25 Juli 2016" Or Form1.DateTimePicker1.Text = "26 Juli 2016" Or Form1.DateTimePicker1.Text = "27 Juli 2016" Or Form1.DateTimePicker1.Text = "28 Juli 2016" Or Form1.DateTimePicker1.Text = "29 Juli 2016" Or Form1.DateTimePicker1.Text = "30 Juli 2016" Or Form1.DateTimePicker1.Text = "31 Juli 2016" Then
            cmd = New OleDbCommand("select * from rplb_juli where NIS = '" & TextBox1.Text & "'", conn)
            dr = cmd.ExecuteReader
            dr.Read()
        ElseIf Form1.DateTimePicker1.Text = "01 Agustus 2016" Or Form1.DateTimePicker1.Text = "02 Agustus 2016" Or Form1.DateTimePicker1.Text = "03 Agustus 2016" Or Form1.DateTimePicker1.Text = "04 Agustus 2016" Or Form1.DateTimePicker1.Text = "05 Agustus 2016" Or Form1.DateTimePicker1.Text = "06 Agustus 2016" Or Form1.DateTimePicker1.Text = "07 Agustus 2016" Or Form1.DateTimePicker1.Text = "08 Agustus 2016" Or Form1.DateTimePicker1.Text = "09 Agustus 2016" Or Form1.DateTimePicker1.Text = "10 Agustus 2016" Or Form1.DateTimePicker1.Text = "11 Agustus 2016" Or Form1.DateTimePicker1.Text = "12 Agustus 2016" Or Form1.DateTimePicker1.Text = "13 Agustus 2016" Or Form1.DateTimePicker1.Text = "14 Agustus 2016" Or Form1.DateTimePicker1.Text = "15 Agustus 2016" Or Form1.DateTimePicker1.Text = "16 Agustus 2016" Or Form1.DateTimePicker1.Text = "17 Agustus 2016" Or Form1.DateTimePicker1.Text = "18 Agustus 2016" Or Form1.DateTimePicker1.Text = "19 Agustus 2016" Or Form1.DateTimePicker1.Text = "20 Agustus 2016" Or Form1.DateTimePicker1.Text = "21 Agustus 2016" Or Form1.DateTimePicker1.Text = "22 Agustus 2016" Or Form1.DateTimePicker1.Text = "23 Agustus 2016" Or Form1.DateTimePicker1.Text = "24 Agustus 2016" Or Form1.DateTimePicker1.Text = "25 Agustus 2016" Or Form1.DateTimePicker1.Text = "26 Agustus 2016" Or Form1.DateTimePicker1.Text = "27 Agustus 2016" Or Form1.DateTimePicker1.Text = "28 Agustus 2016" Or Form1.DateTimePicker1.Text = "29 Agustus 2016" Or Form1.DateTimePicker1.Text = "30 Agustus 2016" Or Form1.DateTimePicker1.Text = "31 Agustus 2016" Then
            cmd = New OleDbCommand("select * from rplb_agustus where NIS = '" & TextBox1.Text & "'", conn)
            dr = cmd.ExecuteReader
            dr.Read()
        ElseIf Form1.DateTimePicker1.Text = "01 September 2016" Or Form1.DateTimePicker1.Text = "02 September 2016" Or Form1.DateTimePicker1.Text = "03 September 2016" Or Form1.DateTimePicker1.Text = "04 September 2016" Or Form1.DateTimePicker1.Text = "05 September 2016" Or Form1.DateTimePicker1.Text = "06 September 2016" Or Form1.DateTimePicker1.Text = "07 September 2016" Or Form1.DateTimePicker1.Text = "08 September 2016" Or Form1.DateTimePicker1.Text = "09 September 2016" Or Form1.DateTimePicker1.Text = "10 September 2016" Or Form1.DateTimePicker1.Text = "11 September 2016" Or Form1.DateTimePicker1.Text = "12 September 2016" Or Form1.DateTimePicker1.Text = "13 September 2016" Or Form1.DateTimePicker1.Text = "14 September 2016" Or Form1.DateTimePicker1.Text = "15 September 2016" Or Form1.DateTimePicker1.Text = "16 September 2016" Or Form1.DateTimePicker1.Text = "17 September 2016" Or Form1.DateTimePicker1.Text = "18 September 2016" Or Form1.DateTimePicker1.Text = "19 September 2016" Or Form1.DateTimePicker1.Text = "20 September 2016" Or Form1.DateTimePicker1.Text = "21 September 2016" Or Form1.DateTimePicker1.Text = "22 September 2016" Or Form1.DateTimePicker1.Text = "23 September 2016" Or Form1.DateTimePicker1.Text = "24 September 2016" Or Form1.DateTimePicker1.Text = "25 September 2016" Or Form1.DateTimePicker1.Text = "26 September 2016" Or Form1.DateTimePicker1.Text = "27 September 2016" Or Form1.DateTimePicker1.Text = "28 September 2016" Or Form1.DateTimePicker1.Text = "29 September 2016" Or Form1.DateTimePicker1.Text = "30 September 2016" Then
            cmd = New OleDbCommand("select * from rplb_september where NIS = '" & TextBox1.Text & "'", conn)
            dr = cmd.ExecuteReader
            dr.Read()
        ElseIf Form1.DateTimePicker1.Text = "01 Oktober 2016" Or Form1.DateTimePicker1.Text = "02 Oktober 2016" Or Form1.DateTimePicker1.Text = "03 Oktober 2016" Or Form1.DateTimePicker1.Text = "04 Oktober 2016" Or Form1.DateTimePicker1.Text = "05 Oktober 2016" Or Form1.DateTimePicker1.Text = "06 Oktober 2016" Or Form1.DateTimePicker1.Text = "07 Oktober 2016" Or Form1.DateTimePicker1.Text = "08 Oktober 2016" Or Form1.DateTimePicker1.Text = "09 Oktober 2016" Or Form1.DateTimePicker1.Text = "10 Oktober 2016" Or Form1.DateTimePicker1.Text = "11 Oktober 2016" Or Form1.DateTimePicker1.Text = "12 Oktober 2016" Or Form1.DateTimePicker1.Text = "13 Oktober 2016" Or Form1.DateTimePicker1.Text = "14 Oktober 2016" Or Form1.DateTimePicker1.Text = "15 Oktober 2016" Or Form1.DateTimePicker1.Text = "16 Oktober 2016" Or Form1.DateTimePicker1.Text = "17 Oktober 2016" Or Form1.DateTimePicker1.Text = "18 Oktober 2016" Or Form1.DateTimePicker1.Text = "19 Oktober 2016" Or Form1.DateTimePicker1.Text = "20 Oktober 2016" Or Form1.DateTimePicker1.Text = "21 Oktober 2016" Or Form1.DateTimePicker1.Text = "22 Oktober 2016" Or Form1.DateTimePicker1.Text = "23 Oktober 2016" Or Form1.DateTimePicker1.Text = "24 Oktober 2016" Or Form1.DateTimePicker1.Text = "25 Oktober 2016" Or Form1.DateTimePicker1.Text = "26 Oktober 2016" Or Form1.DateTimePicker1.Text = "27 Oktober 2016" Or Form1.DateTimePicker1.Text = "28 Oktober 2016" Or Form1.DateTimePicker1.Text = "29 Oktober 2016" Or Form1.DateTimePicker1.Text = "30 Oktober 2016" Or Form1.DateTimePicker1.Text = "31 Oktober 2016" Then
            cmd = New OleDbCommand("select * from rplb_oktober where NIS = '" & TextBox1.Text & "'", conn)
            dr = cmd.ExecuteReader
            dr.Read()
        ElseIf Form1.DateTimePicker1.Text = "01 November 2016" Or Form1.DateTimePicker1.Text = "02 November 2016" Or Form1.DateTimePicker1.Text = "03 November 2016" Or Form1.DateTimePicker1.Text = "04 November 2016" Or Form1.DateTimePicker1.Text = "05 November 2016" Or Form1.DateTimePicker1.Text = "06 November 2016" Or Form1.DateTimePicker1.Text = "07 November 2016" Or Form1.DateTimePicker1.Text = "08 November 2016" Or Form1.DateTimePicker1.Text = "09 November 2016" Or Form1.DateTimePicker1.Text = "10 November 2016" Or Form1.DateTimePicker1.Text = "11 November 2016" Or Form1.DateTimePicker1.Text = "12 November 2016" Or Form1.DateTimePicker1.Text = "13 November 2016" Or Form1.DateTimePicker1.Text = "14 November 2016" Or Form1.DateTimePicker1.Text = "15 November 2016" Or Form1.DateTimePicker1.Text = "16 November 2016" Or Form1.DateTimePicker1.Text = "17 November 2016" Or Form1.DateTimePicker1.Text = "18 November 2016" Or Form1.DateTimePicker1.Text = "19 November 2016" Or Form1.DateTimePicker1.Text = "20 November 2016" Or Form1.DateTimePicker1.Text = "21 November 2016" Or Form1.DateTimePicker1.Text = "22 November 2016" Or Form1.DateTimePicker1.Text = "23 November 2016" Or Form1.DateTimePicker1.Text = "24 November 2016" Or Form1.DateTimePicker1.Text = "25 November 2016" Or Form1.DateTimePicker1.Text = "26 November 2016" Or Form1.DateTimePicker1.Text = "27 November 2016" Or Form1.DateTimePicker1.Text = "28 November 2016" Or Form1.DateTimePicker1.Text = "29 November 2016" Or Form1.DateTimePicker1.Text = "30 November 2016" Then
            cmd = New OleDbCommand("select * from rplb_november where NIS = '" & TextBox1.Text & "'", conn)
            dr = cmd.ExecuteReader
            dr.Read()
        ElseIf Form1.DateTimePicker1.Text = "01 Desember 2016" Or Form1.DateTimePicker1.Text = "02 Desember 2016" Or Form1.DateTimePicker1.Text = "03 Desember 2016" Or Form1.DateTimePicker1.Text = "04 Desember 2016" Or Form1.DateTimePicker1.Text = "05 Desember 2016" Or Form1.DateTimePicker1.Text = "06 Desember 2016" Or Form1.DateTimePicker1.Text = "07 Desember 2016" Or Form1.DateTimePicker1.Text = "08 Desember 2016" Or Form1.DateTimePicker1.Text = "09 Desember 2016" Or Form1.DateTimePicker1.Text = "10 Desember 2016" Or Form1.DateTimePicker1.Text = "11 Desember 2016" Or Form1.DateTimePicker1.Text = "12 Desember 2016" Or Form1.DateTimePicker1.Text = "13 Desember 2016" Or Form1.DateTimePicker1.Text = "14 Desember 2016" Or Form1.DateTimePicker1.Text = "15 Desember 2016" Or Form1.DateTimePicker1.Text = "16 Desember 2016" Or Form1.DateTimePicker1.Text = "17 Desember 2016" Or Form1.DateTimePicker1.Text = "18 Desember 2016" Or Form1.DateTimePicker1.Text = "19 Desember 2016" Or Form1.DateTimePicker1.Text = "20 Desember 2016" Or Form1.DateTimePicker1.Text = "21 Desember 2016" Or Form1.DateTimePicker1.Text = "22 Desember 2016" Or Form1.DateTimePicker1.Text = "23 Desember 2016" Or Form1.DateTimePicker1.Text = "24 Desember 2016" Or Form1.DateTimePicker1.Text = "25 Desember 2016" Or Form1.DateTimePicker1.Text = "26 Desember 2016" Or Form1.DateTimePicker1.Text = "27 Desember 2016" Or Form1.DateTimePicker1.Text = "28 Desember 2016" Or Form1.DateTimePicker1.Text = "29 Desember 2016" Or Form1.DateTimePicker1.Text = "30 Desember 2016" Or Form1.DateTimePicker1.Text = "31 Desember 2016" Then
            cmd = New OleDbCommand("select * from rplb_desember where NIS = '" & TextBox1.Text & "'", conn)
            dr = cmd.ExecuteReader
            dr.Read()
        ElseIf Form1.DateTimePicker1.Text = "01 Januari 2017" Or Form1.DateTimePicker1.Text = "02 Januari 2017" Or Form1.DateTimePicker1.Text = "03 Januari 2017" Or Form1.DateTimePicker1.Text = "04 Januari 2017" Or Form1.DateTimePicker1.Text = "05 Januari 2017" Or Form1.DateTimePicker1.Text = "06 Januari 2017" Or Form1.DateTimePicker1.Text = "07 Januari 2017" Or Form1.DateTimePicker1.Text = "08 Januari 2017" Or Form1.DateTimePicker1.Text = "09 Januari 2017" Or Form1.DateTimePicker1.Text = "10 Januari 2017" Or Form1.DateTimePicker1.Text = "11 Januari 2017" Or Form1.DateTimePicker1.Text = "12 Januari 2017" Or Form1.DateTimePicker1.Text = "13 Januari 2017" Or Form1.DateTimePicker1.Text = "14 Januari 2017" Or Form1.DateTimePicker1.Text = "15 Januari 2017" Or Form1.DateTimePicker1.Text = "16 Januari 2017" Or Form1.DateTimePicker1.Text = "17 Januari 2017" Or Form1.DateTimePicker1.Text = "18 Januari 2017" Or Form1.DateTimePicker1.Text = "19 Januari 2017" Or Form1.DateTimePicker1.Text = "20 Januari 2017" Or Form1.DateTimePicker1.Text = "21 Januari 2017" Or Form1.DateTimePicker1.Text = "22 Januari 2017" Or Form1.DateTimePicker1.Text = "23 Januari 2017" Or Form1.DateTimePicker1.Text = "24 Januari 2017" Or Form1.DateTimePicker1.Text = "25 Januari 2017" Or Form1.DateTimePicker1.Text = "26 Januari 2017" Or Form1.DateTimePicker1.Text = "27 Januari 2017" Or Form1.DateTimePicker1.Text = "28 Januari 2017" Or Form1.DateTimePicker1.Text = "29 Januari 2017" Or Form1.DateTimePicker1.Text = "30 Januari 2017" Or Form1.DateTimePicker1.Text = "31 Januari 2017" Then
            cmd = New OleDbCommand("select * from rplb_januari where NIS = '" & TextBox1.Text & "'", conn)
            dr = cmd.ExecuteReader
            dr.Read()
        ElseIf Form1.DateTimePicker1.Text = "01 Februari 2017" Or Form1.DateTimePicker1.Text = "02 Februari 2017" Or Form1.DateTimePicker1.Text = "03 Februari 2017" Or Form1.DateTimePicker1.Text = "04 Februari 2017" Or Form1.DateTimePicker1.Text = "05 Februari 2017" Or Form1.DateTimePicker1.Text = "06 Februari 2017" Or Form1.DateTimePicker1.Text = "07 Februari 2017" Or Form1.DateTimePicker1.Text = "08 Februari 2017" Or Form1.DateTimePicker1.Text = "09 Februari 2017" Or Form1.DateTimePicker1.Text = "10 Februari 2017" Or Form1.DateTimePicker1.Text = "11 Februari 2017" Or Form1.DateTimePicker1.Text = "12 Februari 2017" Or Form1.DateTimePicker1.Text = "13 Februari 2017" Or Form1.DateTimePicker1.Text = "14 Februari 2017" Or Form1.DateTimePicker1.Text = "15 Februari 2017" Or Form1.DateTimePicker1.Text = "16 Februari 2017" Or Form1.DateTimePicker1.Text = "17 Februari 2017" Or Form1.DateTimePicker1.Text = "18 Februari 2017" Or Form1.DateTimePicker1.Text = "19 Februari 2017" Or Form1.DateTimePicker1.Text = "20 Februari 2017" Or Form1.DateTimePicker1.Text = "21 Februari 2017" Or Form1.DateTimePicker1.Text = "22 Februari 2017" Or Form1.DateTimePicker1.Text = "23 Februari 2017" Or Form1.DateTimePicker1.Text = "24 Februari 2017" Or Form1.DateTimePicker1.Text = "25 Februari 2017" Or Form1.DateTimePicker1.Text = "26 Februari 2017" Or Form1.DateTimePicker1.Text = "27 Februari 2017" Or Form1.DateTimePicker1.Text = "28 Februari 2017" Then
            cmd = New OleDbCommand("select * from rplb_februari where NIS = '" & TextBox1.Text & "'", conn)
            dr = cmd.ExecuteReader
            dr.Read()
        ElseIf Form1.DateTimePicker1.Text = "01 Maret 2017" Or Form1.DateTimePicker1.Text = "02 Maret 2017" Or Form1.DateTimePicker1.Text = "03 Maret 2017" Or Form1.DateTimePicker1.Text = "04 Maret 2017" Or Form1.DateTimePicker1.Text = "05 Maret 2017" Or Form1.DateTimePicker1.Text = "06 Maret 2017" Or Form1.DateTimePicker1.Text = "07 Maret 2017" Or Form1.DateTimePicker1.Text = "08 Maret 2017" Or Form1.DateTimePicker1.Text = "09 Maret 2017" Or Form1.DateTimePicker1.Text = "10 Maret 2017" Or Form1.DateTimePicker1.Text = "11 Maret 2017" Or Form1.DateTimePicker1.Text = "12 Maret 2017" Or Form1.DateTimePicker1.Text = "13 Maret 2017" Or Form1.DateTimePicker1.Text = "14 Maret 2017" Or Form1.DateTimePicker1.Text = "15 Maret 2017" Or Form1.DateTimePicker1.Text = "16 Maret 2017" Or Form1.DateTimePicker1.Text = "17 Maret 2017" Or Form1.DateTimePicker1.Text = "18 Maret 2017" Or Form1.DateTimePicker1.Text = "19 Maret 2017" Or Form1.DateTimePicker1.Text = "20 Maret 2017" Or Form1.DateTimePicker1.Text = "21 Maret 2017" Or Form1.DateTimePicker1.Text = "22 Maret 2017" Or Form1.DateTimePicker1.Text = "23 Maret 2017" Or Form1.DateTimePicker1.Text = "24 Maret 2017" Or Form1.DateTimePicker1.Text = "25 Maret 2017" Or Form1.DateTimePicker1.Text = "26 Maret 2017" Or Form1.DateTimePicker1.Text = "27 Maret 2017" Or Form1.DateTimePicker1.Text = "28 Maret 2017" Or Form1.DateTimePicker1.Text = "29 Maret 2017" Or Form1.DateTimePicker1.Text = "30 Maret 2017" Or Form1.DateTimePicker1.Text = "31 Maret 2017" Then
            cmd = New OleDbCommand("select * from rplb_maret where NIS = '" & TextBox1.Text & "'", conn)
            dr = cmd.ExecuteReader
            dr.Read()
        ElseIf Form1.DateTimePicker1.Text = "01 April 2017" Or Form1.DateTimePicker1.Text = "02 April 2017" Or Form1.DateTimePicker1.Text = "03 April 2017" Or Form1.DateTimePicker1.Text = "04 April 2017" Or Form1.DateTimePicker1.Text = "05 April 2017" Or Form1.DateTimePicker1.Text = "06 April 2017" Or Form1.DateTimePicker1.Text = "07 April 2017" Or Form1.DateTimePicker1.Text = "08 April 2017" Or Form1.DateTimePicker1.Text = "09 April 2017" Or Form1.DateTimePicker1.Text = "10 April 2017" Or Form1.DateTimePicker1.Text = "11 April 2017" Or Form1.DateTimePicker1.Text = "12 April 2017" Or Form1.DateTimePicker1.Text = "13 April 2017" Or Form1.DateTimePicker1.Text = "14 April 2017" Or Form1.DateTimePicker1.Text = "15 April 2017" Or Form1.DateTimePicker1.Text = "16 April 2017" Or Form1.DateTimePicker1.Text = "17 April 2017" Or Form1.DateTimePicker1.Text = "18 April 2017" Or Form1.DateTimePicker1.Text = "19 April 2017" Or Form1.DateTimePicker1.Text = "20 April 2017" Or Form1.DateTimePicker1.Text = "21 April 2017" Or Form1.DateTimePicker1.Text = "22 April 2017" Or Form1.DateTimePicker1.Text = "23 April 2017" Or Form1.DateTimePicker1.Text = "24 April 2017" Or Form1.DateTimePicker1.Text = "25 April 2017" Or Form1.DateTimePicker1.Text = "26 April 2017" Or Form1.DateTimePicker1.Text = "27 April 2017" Or Form1.DateTimePicker1.Text = "28 April 2017" Or Form1.DateTimePicker1.Text = "29 April 2017" Or Form1.DateTimePicker1.Text = "30 April 2017" Then
            cmd = New OleDbCommand("select * from rplb_april where NIS = '" & TextBox1.Text & "'", conn)
            dr = cmd.ExecuteReader
            dr.Read()
        ElseIf Form1.DateTimePicker1.Text = "01 Mei 2017" Or Form1.DateTimePicker1.Text = "02 Mei 2017" Or Form1.DateTimePicker1.Text = "03 Mei 2017" Or Form1.DateTimePicker1.Text = "04 Mei 2017" Or Form1.DateTimePicker1.Text = "05 Mei 2017" Or Form1.DateTimePicker1.Text = "06 Mei 2017" Or Form1.DateTimePicker1.Text = "07 Mei 2017" Or Form1.DateTimePicker1.Text = "08 Mei 2017" Or Form1.DateTimePicker1.Text = "09 Mei 2017" Or Form1.DateTimePicker1.Text = "10 Mei 2017" Or Form1.DateTimePicker1.Text = "11 Mei 2017" Or Form1.DateTimePicker1.Text = "12 Mei 2017" Or Form1.DateTimePicker1.Text = "13 Mei 2017" Or Form1.DateTimePicker1.Text = "14 Mei 2017" Or Form1.DateTimePicker1.Text = "15 Mei 2017" Or Form1.DateTimePicker1.Text = "16 Mei 2017" Or Form1.DateTimePicker1.Text = "17 Mei 2017" Or Form1.DateTimePicker1.Text = "18 Mei 2017" Or Form1.DateTimePicker1.Text = "19 Mei 2017" Or Form1.DateTimePicker1.Text = "20 Mei 2017" Or Form1.DateTimePicker1.Text = "21 Mei 2017" Or Form1.DateTimePicker1.Text = "22 Mei 2017" Or Form1.DateTimePicker1.Text = "23 Mei 2017" Or Form1.DateTimePicker1.Text = "24 Mei 2017" Or Form1.DateTimePicker1.Text = "25 Mei 2017" Or Form1.DateTimePicker1.Text = "26 Mei 2017" Or Form1.DateTimePicker1.Text = "27 Mei 2017" Or Form1.DateTimePicker1.Text = "28 Mei 2017" Or Form1.DateTimePicker1.Text = "29 Mei 2017" Or Form1.DateTimePicker1.Text = "30 Mei 2017" Or Form1.DateTimePicker1.Text = "31 Mei 2017" Then
            cmd = New OleDbCommand("select * from rplb_mei where NIS = '" & TextBox1.Text & "'", conn)
            dr = cmd.ExecuteReader
            dr.Read()
        ElseIf Form1.DateTimePicker1.Text = "01 Juni 2017" Or Form1.DateTimePicker1.Text = "02 Juni 2017" Or Form1.DateTimePicker1.Text = "03 Juni 2017" Or Form1.DateTimePicker1.Text = "04 Juni 2017" Or Form1.DateTimePicker1.Text = "05 Juni 2017" Or Form1.DateTimePicker1.Text = "06 Juni 2017" Or Form1.DateTimePicker1.Text = "07 Juni 2017" Or Form1.DateTimePicker1.Text = "08 Juni 2017" Or Form1.DateTimePicker1.Text = "09 Juni 2017" Or Form1.DateTimePicker1.Text = "10 Juni 2017" Or Form1.DateTimePicker1.Text = "11 Juni 2017" Or Form1.DateTimePicker1.Text = "12 Juni 2017" Or Form1.DateTimePicker1.Text = "13 Juni 2017" Or Form1.DateTimePicker1.Text = "14 Juni 2017" Or Form1.DateTimePicker1.Text = "15 Juni 2017" Or Form1.DateTimePicker1.Text = "16 Juni 2017" Or Form1.DateTimePicker1.Text = "17 Juni 2017" Or Form1.DateTimePicker1.Text = "18 Juni 2017" Or Form1.DateTimePicker1.Text = "19 Juni 2017" Or Form1.DateTimePicker1.Text = "20 Juni 2017" Or Form1.DateTimePicker1.Text = "21 Juni 2017" Or Form1.DateTimePicker1.Text = "22 Juni 2017" Or Form1.DateTimePicker1.Text = "23 Juni 2017" Or Form1.DateTimePicker1.Text = "24 Juni 2017" Or Form1.DateTimePicker1.Text = "25 Juni 2017" Or Form1.DateTimePicker1.Text = "26 Juni 2017" Or Form1.DateTimePicker1.Text = "27 Juni 2017" Or Form1.DateTimePicker1.Text = "28 Juni 2017" Or Form1.DateTimePicker1.Text = "29 Juni 2017" Or Form1.DateTimePicker1.Text = "30 Juni 2017" Then
            cmd = New OleDbCommand("select * from rplb_juni where NIS = '" & TextBox1.Text & "'", conn)
            dr = cmd.ExecuteReader
            dr.Read()
        End If
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Try

            If Form1.DateTimePicker1.Text = "01 Juli 2016" Then
                Call koneksi()
                Dim edit As String = "update rplb_juli set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "02 Juli 2016" Then
                Call koneksi()
                Dim edit As String = "update rplb_juli set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "03 Juli 2016" Then
                Call koneksi()
                Dim edit As String = "update rplb_juli set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "04 Juli 2016" Then
                Call koneksi()
                Dim edit As String = "update rplb_juli set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "05 Juli 2016" Then
                Call koneksi()
                Dim edit As String = "update rplb_juli set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "06 Juli 2016" Then
                Call koneksi()
                Dim edit As String = "update rplb_juli set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "07 Juli 2016" Then
                Call koneksi()
                Dim edit As String = "update rplb_juli set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "09 Juli 2016" Then
                Call koneksi()
                Dim edit As String = "update rplb_juli set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "10 Juli 2016" Then
                Call koneksi()
                Dim edit As String = "update rplb_juli set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "11 Juli 2016" Then
                Call koneksi()
                Dim edit As String = "update rplb_juli set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "12 Juli 2016" Then
                Call koneksi()
                Dim edit As String = "update rplb_juli set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "13 Juli 2016" Then
                Call koneksi()
                Dim edit As String = "update rplb_juli set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "14 Juli 2016" Then
                Call koneksi()
                Dim edit As String = "update rplb_juli set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "15 Juli 2016" Then
                Call koneksi()
                Dim edit As String = "update rplb_juli set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "16 Juli 2016" Then
                Call koneksi()
                Dim edit As String = "update rplb_juli set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "17 Juli 2016" Then
                Call koneksi()
                Dim edit As String = "update rplb_juli set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "18 Juli 2016" Then
                Call koneksi()
                Dim edit As String = "update rplb_juli set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "19 Juli 2016" Then
                Call koneksi()
                Dim edit As String = "update rplb_juli set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "20 Juli 2016" Then
                Call koneksi()
                Dim edit As String = "update rplb_juli set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "21 Juli 2016" Then
                Call koneksi()
                Dim edit As String = "update rplb_juli set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "22 Juli 2016" Then
                Call koneksi()
                Dim edit As String = "update rplb_juli set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "23 Juli 2016" Then
                Call koneksi()
                Dim edit As String = "update rplb_juli set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "24 Juli 2016" Then
                Call koneksi()
                Dim edit As String = "update rplb_juli set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "25 Juli 2016" Then
                Call koneksi()
                Dim edit As String = "update rplb_juli set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "26 Juli 2016" Then
                Call koneksi()
                Dim edit As String = "update rplb_juli set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "27 Juli 2016" Then
                Call koneksi()
                Dim edit As String = "update rplb_juli set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "28 Juli 2016" Then
                Call koneksi()
                Dim edit As String = "update rplb_juli set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "29 Juli 2016" Then
                Call koneksi()
                Dim edit As String = "update rplb_juli set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "30 Juli 2016" Then
                Call koneksi()
                Dim edit As String = "update rplb_juli set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "31 Juli 2016" Then
                Call koneksi()
                Dim edit As String = "update rplb_juli set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()



            ElseIf Form1.DateTimePicker1.Text = "01 Agustus 2016" Then
                Call koneksi()
                Dim edit As String = "update rplb_agustus set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "02 Agustus 2016" Then
                Call koneksi()
                Dim edit As String = "update rplb_agustus set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "03 Agustus 2016" Then
                Call koneksi()
                Dim edit As String = "update rplb_agustus set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "04 Agustus 2016" Then
                Call koneksi()
                Dim edit As String = "update rplb_agustus set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "05 Agustus 2016" Then
                Call koneksi()
                Dim edit As String = "update rplb_agustus set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "06 Agustus 2016" Then
                Call koneksi()
                Dim edit As String = "update rplb_agustus set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "07 Agustus 2016" Then
                Call koneksi()
                Dim edit As String = "update rplb_agustus set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "08 Agustus 2016" Then
                Call koneksi()
                Dim edit As String = "update rplb_agustus set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "09 Agustus 2016" Then
                Call koneksi()
                Dim edit As String = "update rplb_agustus set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "10 Agustus 2016" Then
                Call koneksi()
                Dim edit As String = "update rplb_agustus set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "11 Agustus 2016" Then
                Call koneksi()
                Dim edit As String = "update rplb_agustus set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "12 Agustus 2016" Then
                Call koneksi()
                Dim edit As String = "update rplb_agustus set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "13 Agustus 2016" Then
                Call koneksi()
                Dim edit As String = "update rplb_agustus set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "14 Agustus 2016" Then
                Call koneksi()
                Dim edit As String = "update rplb_agustus set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "15 Agustus 2016" Then
                Call koneksi()
                Dim edit As String = "update rplb_agustus set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "16 Agustus 2016" Then
                Call koneksi()
                Dim edit As String = "update rplb_agustus set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "17 Agustus 2016" Then
                Call koneksi()
                Dim edit As String = "update rplb_agustus set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "18 Agustus 2016" Then
                Call koneksi()
                Dim edit As String = "update rplb_agustus set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "19 Agustus 2016" Then
                Call koneksi()
                Dim edit As String = "update rplb_agustus set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "20 Agustus 2016" Then
                Call koneksi()
                Dim edit As String = "update rplb_agustus set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "21 Agustus 2016" Then
                Call koneksi()
                Dim edit As String = "update rplb_agustus set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "22 Agustus 2016" Then
                Call koneksi()
                Dim edit As String = "update rplb_agustus set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "23 Agustus 2016" Then
                Call koneksi()
                Dim edit As String = "update rplb_agustus set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "24 Agustus 2016" Then
                Call koneksi()
                Dim edit As String = "update rplb_agustus set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "25 Agustus 2016" Then
                Call koneksi()
                Dim edit As String = "update rplb_agustus set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "26 Agustus 2016" Then
                Call koneksi()
                Dim edit As String = "update rplb_agustus set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "27 Agustus 2016" Then
                Call koneksi()
                Dim edit As String = "update rplb_agustus set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "28 Agustus 2016" Then
                Call koneksi()
                Dim edit As String = "update rplb_agustus set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "29 Agustus 2016" Then
                Call koneksi()
                Dim edit As String = "update rplb_agustus set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "30 Agustus 2016" Then
                Call koneksi()
                Dim edit As String = "update rplb_agustus set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "31 Agustus 2016" Then
                Call koneksi()
                Dim edit As String = "update rplb_agustus set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()


            ElseIf Form1.DateTimePicker1.Text = "01 September 2016" Then
                Call koneksi()
                Dim edit As String = "update rplb_september set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "02 September 2016" Then
                Call koneksi()
                Dim edit As String = "update rplb_september set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "03 September 2016" Then
                Call koneksi()
                Dim edit As String = "update rplb_september set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "04 September 2016" Then
                Call koneksi()
                Dim edit As String = "update rplb_september set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "05 September 2016" Then
                Call koneksi()
                Dim edit As String = "update rplb_september set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "06 September 2016" Then
                Call koneksi()
                Dim edit As String = "update rplb_september set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "07 September 2016" Then
                Call koneksi()
                Dim edit As String = "update rplb_september set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "08 September 2016" Then
                Call koneksi()
                Dim edit As String = "update rplb_september set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "09 September 2016" Then
                Call koneksi()
                Dim edit As String = "update rplb_september set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "10 September 2016" Then
                Call koneksi()
                Dim edit As String = "update rplb_september set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "11 September 2016" Then
                Call koneksi()
                Dim edit As String = "update rplb_september set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "12 September 2016" Then
                Call koneksi()
                Dim edit As String = "update rplb_september set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "13 September 2016" Then
                Call koneksi()
                Dim edit As String = "update rplb_september set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "14 September 2016" Then
                Call koneksi()
                Dim edit As String = "update rplb_september set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "15 September 2016" Then
                Call koneksi()
                Dim edit As String = "update rplb_september set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "16 September 2016" Then
                Call koneksi()
                Dim edit As String = "update rplb_september set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "17 September 2016" Then
                Call koneksi()
                Dim edit As String = "update rplb_september set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "18 September 2016" Then
                Call koneksi()
                Dim edit As String = "update rplb_september set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "19 September 2016" Then
                Call koneksi()
                Dim edit As String = "update rplb_september set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "20 September 2016" Then
                Call koneksi()
                Dim edit As String = "update rplb_september set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "21 September 2016" Then
                Call koneksi()
                Dim edit As String = "update rplb_september set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "22 September 2016" Then
                Call koneksi()
                Dim edit As String = "update rplb_september set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "23 September 2016" Then
                Call koneksi()
                Dim edit As String = "update rplb_september set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "24 September 2016" Then
                Call koneksi()
                Dim edit As String = "update rplb_september set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "25 September 2016" Then
                Call koneksi()
                Dim edit As String = "update rplb_september set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "26 September 2016" Then
                Call koneksi()
                Dim edit As String = "update rplb_september set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "27 September 2016" Then
                Call koneksi()
                Dim edit As String = "update rplb_september set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "28 September 2016" Then
                Call koneksi()
                Dim edit As String = "update rplb_september set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "29 September 2016" Then
                Call koneksi()
                Dim edit As String = "update rplb_september set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "30 September 2016" Then
                Call koneksi()
                Dim edit As String = "update rplb_september set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()



            ElseIf Form1.DateTimePicker1.Text = "01 Oktober 2016" Then
                Call koneksi()
                Dim edit As String = "update rplb_oktober set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "02 Oktober 2016" Then
                Call koneksi()
                Dim edit As String = "update rplb_oktober set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "03 Oktober 2016" Then
                Call koneksi()
                Dim edit As String = "update rplb_oktober set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "04 Oktober 2016" Then
                Call koneksi()
                Dim edit As String = "update rplb_oktober set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "05 Oktober 2016" Then
                Call koneksi()
                Dim edit As String = "update rplb_oktober set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "06 Oktober 2016" Then
                Call koneksi()
                Dim edit As String = "update rplb_oktober set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "07 Oktober 2016" Then
                Call koneksi()
                Dim edit As String = "update rplb_oktober set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "08 Oktober 2016" Then
                Call koneksi()
                Dim edit As String = "update rplb_oktober set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "09 Oktober 2016" Then
                Call koneksi()
                Dim edit As String = "update rplb_oktober set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "10 Oktober 2016" Then
                Call koneksi()
                Dim edit As String = "update rplb_oktober set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "11 Oktober 2016" Then
                Call koneksi()
                Dim edit As String = "update rplb_oktober set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "12 Oktober 2016" Then
                Call koneksi()
                Dim edit As String = "update rplb_oktober set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "13 Oktober 2016" Then
                Call koneksi()
                Dim edit As String = "update rplb_oktober set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "14 Oktober 2016" Then
                Call koneksi()
                Dim edit As String = "update rplb_oktober set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "15 Oktober 2016" Then
                Call koneksi()
                Dim edit As String = "update rplb_oktober set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "16 Oktober 2016" Then
                Call koneksi()
                Dim edit As String = "update rplb_oktober set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "17 Oktober 2016" Then
                Call koneksi()
                Dim edit As String = "update rplb_oktober set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "18 Oktober 2016" Then
                Call koneksi()
                Dim edit As String = "update rplb_oktober set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "19 Oktober 2016" Then
                Call koneksi()
                Dim edit As String = "update rplb_oktober set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "20 Oktober 2016" Then
                Call koneksi()
                Dim edit As String = "update rplb_oktober set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "21 Oktober 2016" Then
                Call koneksi()
                Dim edit As String = "update rplb_oktober set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "22 Oktober 2016" Then
                Call koneksi()
                Dim edit As String = "update rplb_oktober set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "23 Oktober 2016" Then
                Call koneksi()
                Dim edit As String = "update rplb_oktober set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "24 Oktober 2016" Then
                Call koneksi()
                Dim edit As String = "update rplb_oktober set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "25 Oktober 2016" Then
                Call koneksi()
                Dim edit As String = "update rplb_oktober set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "26 Oktober 2016" Then
                Call koneksi()
                Dim edit As String = "update rplb_oktober set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "27 Oktober 2016" Then
                Call koneksi()
                Dim edit As String = "update rplb_oktober set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "28 Oktober 2016" Then
                Call koneksi()
                Dim edit As String = "update rplb_oktober set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "29 Oktober 2016" Then
                Call koneksi()
                Dim edit As String = "update rplb_oktober set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "30 Oktober 2016" Then
                Call koneksi()
                Dim edit As String = "update rplb_oktober set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "31 Oktober 2016" Then
                Call koneksi()
                Dim edit As String = "update rplb_oktober set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()


            ElseIf Form1.DateTimePicker1.Text = "01 November 2016" Then
                Call koneksi()
                Dim edit As String = "update rplb_november set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "02 November 2016" Then
                Call koneksi()
                Dim edit As String = "update rplb_november set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "03 November 2016" Then
                Call koneksi()
                Dim edit As String = "update rplb_november set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "04 November 2016" Then
                Call koneksi()
                Dim edit As String = "update rplb_november set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "05 November 2016" Then
                Call koneksi()
                Dim edit As String = "update rplb_november set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "06 November 2016" Then
                Call koneksi()
                Dim edit As String = "update rplb_november set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "07 November 2016" Then
                Call koneksi()
                Dim edit As String = "update rplb_november set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "08 November 2016" Then
                Call koneksi()
                Dim edit As String = "update rplb_november set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "09 November 2016" Then
                Call koneksi()
                Dim edit As String = "update rplb_november set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "10 November 2016" Then
                Call koneksi()
                Dim edit As String = "update rplb_november set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "11 November 2016" Then
                Call koneksi()
                Dim edit As String = "update rplb_november set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "12 November 2016" Then
                Call koneksi()
                Dim edit As String = "update rplb_november set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "13 November 2016" Then
                Call koneksi()
                Dim edit As String = "update rplb_november set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "14 November 2016" Then
                Call koneksi()
                Dim edit As String = "update rplb_november set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "15 November 2016" Then
                Call koneksi()
                Dim edit As String = "update rplb_november set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "16 November 2016" Then
                Call koneksi()
                Dim edit As String = "update rplb_november set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "17 November 2016" Then
                Call koneksi()
                Dim edit As String = "update rplb_november set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "18 November 2016" Then
                Call koneksi()
                Dim edit As String = "update rplb_november set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "19 November 2016" Then
                Call koneksi()
                Dim edit As String = "update rplb_november set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "20 November 2016" Then
                Call koneksi()
                Dim edit As String = "update rplb_november set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "21 November 2016" Then
                Call koneksi()
                Dim edit As String = "update rplb_november set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "22 November 2016" Then
                Call koneksi()
                Dim edit As String = "update rplb_november set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "23 November 2016" Then
                Call koneksi()
                Dim edit As String = "update rplb_november set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "24 November 2016" Then
                Call koneksi()
                Dim edit As String = "update rplb_november set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "25 November 2016" Then
                Call koneksi()
                Dim edit As String = "update rplb_november set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "26 November 2016" Then
                Call koneksi()
                Dim edit As String = "update rplb_november set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "27 November 2016" Then
                Call koneksi()
                Dim edit As String = "update rplb_november set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "28 November 2016" Then
                Call koneksi()
                Dim edit As String = "update rplb_november set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "29 November 2016" Then
                Call koneksi()
                Dim edit As String = "update rplb_november set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "30 November 2016" Then
                Call koneksi()
                Dim edit As String = "update rplb_november set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()


            ElseIf Form1.DateTimePicker1.Text = "01 Desember 2016" Then
                Call koneksi()
                Dim edit As String = "update rplb_desember set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "02 Desember 2016" Then
                Call koneksi()
                Dim edit As String = "update rplb_desember set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "03 Desember 2016" Then
                Call koneksi()
                Dim edit As String = "update rplb_desember set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "04 Desember 2016" Then
                Call koneksi()
                Dim edit As String = "update rplb_desember set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "05 Desember 2016" Then
                Call koneksi()
                Dim edit As String = "update rplb_desember set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "06 Desember 2016" Then
                Call koneksi()
                Dim edit As String = "update rplb_desember set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "07 Desember 2016" Then
                Call koneksi()
                Dim edit As String = "update rplb_desember set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "08 Desember 2016" Then
                Call koneksi()
                Dim edit As String = "update rplb_desember set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "09 Desember 2016" Then
                Call koneksi()
                Dim edit As String = "update rplb_desember set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "10 Desember 2016" Then
                Call koneksi()
                Dim edit As String = "update rplb_desember set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "11 Desember 2016" Then
                Call koneksi()
                Dim edit As String = "update rplb_desember set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "12 Desember 2016" Then
                Call koneksi()
                Dim edit As String = "update rplb_desember set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "13 Desember 2016" Then
                Call koneksi()
                Dim edit As String = "update rplb_desember set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "14 Desember 2016" Then
                Call koneksi()
                Dim edit As String = "update rplb_desember set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "15 Desember 2016" Then
                Call koneksi()
                Dim edit As String = "update rplb_desember set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "16 Desember 2016" Then
                Call koneksi()
                Dim edit As String = "update rplb_desember set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "17 Desember 2016" Then
                Call koneksi()
                Dim edit As String = "update rplb_desember set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "18 Desember 2016" Then
                Call koneksi()
                Dim edit As String = "update rplb_desember set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "19 Desember 2016" Then
                Call koneksi()
                Dim edit As String = "update rplb_desember set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "20 Desember 2016" Then
                Call koneksi()
                Dim edit As String = "update rplb_desember set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "21 Desember 2016" Then
                Call koneksi()
                Dim edit As String = "update rplb_desember set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "22 Desember 2016" Then
                Call koneksi()
                Dim edit As String = "update rplb_desember set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "23 Desember 2016" Then
                Call koneksi()
                Dim edit As String = "update rplb_desember set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "24 Desember 2016" Then
                Call koneksi()
                Dim edit As String = "update rplb_desember set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "25 Desember 2016" Then
                Call koneksi()
                Dim edit As String = "update rplb_desember set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "26 Desember 2016" Then
                Call koneksi()
                Dim edit As String = "update rplb_desember set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "27 Desember 2016" Then
                Call koneksi()
                Dim edit As String = "update rplb_desember set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "28 Desember 2016" Then
                Call koneksi()
                Dim edit As String = "update rplb_desember set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "29 Desember 2016" Then
                Call koneksi()
                Dim edit As String = "update rplb_desember set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "30 Desember 2016" Then
                Call koneksi()
                Dim edit As String = "update rplb_desember set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "31 Desember 2016" Then
                Call koneksi()
                Dim edit As String = "update rplb_desember set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()


            ElseIf Form1.DateTimePicker1.Text = "01 Januari 2017" Then
                Call koneksi()
                Dim edit As String = "update rplb_januari set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "02 Januari 2017" Then
                Call koneksi()
                Dim edit As String = "update rplb_januari set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "03 Januari 2017" Then
                Call koneksi()
                Dim edit As String = "update rplb_januari set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "04 Januari 2017" Then
                Call koneksi()
                Dim edit As String = "update rplb_januari set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "05 Januari 2017" Then
                Call koneksi()
                Dim edit As String = "update rplb_januari set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "06 Januari 2017" Then
                Call koneksi()
                Dim edit As String = "update rplb_januari set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "07 Januari 2017" Then
                Call koneksi()
                Dim edit As String = "update rplb_januari set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "08 Januari 2017" Then
                Call koneksi()
                Dim edit As String = "update rplb_januari set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "09 Januari 2017" Then
                Call koneksi()
                Dim edit As String = "update rplb_januari set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "10 Januari 2017" Then
                Call koneksi()
                Dim edit As String = "update rplb_januari set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "11 Januari 2017" Then
                Call koneksi()
                Dim edit As String = "update rplb_januari set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "12 Januari 2017" Then
                Call koneksi()
                Dim edit As String = "update rplb_januari set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "13 Januari 2017" Then
                Call koneksi()
                Dim edit As String = "update rplb_januari set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "14 Januari 2017" Then
                Call koneksi()
                Dim edit As String = "update rplb_januari set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "15 Januari 2017" Then
                Call koneksi()
                Dim edit As String = "update rplb_januari set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "16 Januari 2017" Then
                Call koneksi()
                Dim edit As String = "update rplb_januari set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "17 Januari 2017" Then
                Call koneksi()
                Dim edit As String = "update rplb_januari set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "18 Januari 2017" Then
                Call koneksi()
                Dim edit As String = "update rplb_januari set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "19 Januari 2017" Then
                Call koneksi()
                Dim edit As String = "update rplb_januari set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "20 Januari 2017" Then
                Call koneksi()
                Dim edit As String = "update rplb_januari set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "21 Januari 2017" Then
                Call koneksi()
                Dim edit As String = "update rplb_januari set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "22 Januari 2017" Then
                Call koneksi()
                Dim edit As String = "update rplb_januari set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "23 Januari 2017" Then
                Call koneksi()
                Dim edit As String = "update rplb_januari set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "24 Januari 2017" Then
                Call koneksi()
                Dim edit As String = "update rplb_januari set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "25 Januari 2017" Then
                Call koneksi()
                Dim edit As String = "update rplb_januari set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "26 Januari 2017" Then
                Call koneksi()
                Dim edit As String = "update rplb_januari set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "27 Januari 2017" Then
                Call koneksi()
                Dim edit As String = "update rplb_januari set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "28 Januari 2017" Then
                Call koneksi()
                Dim edit As String = "update rplb_januari set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "29 Januari 2017" Then
                Call koneksi()
                Dim edit As String = "update rplb_januari set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "30 Januari 2017" Then
                Call koneksi()
                Dim edit As String = "update rplb_januari set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "31 Januari 2017" Then
                Call koneksi()
                Dim edit As String = "update rplb_januari set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()


            ElseIf Form1.DateTimePicker1.Text = "01 Februari 2017" Then
                Call koneksi()
                Dim edit As String = "update rplb_februari set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "02 Februari 2017" Then
                Call koneksi()
                Dim edit As String = "update rplb_februari set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "03 Februari 2017" Then
                Call koneksi()
                Dim edit As String = "update rplb_februari set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "04 Februari 2017" Then
                Call koneksi()
                Dim edit As String = "update rplb_februari set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "05 Februari 2017" Then
                Call koneksi()
                Dim edit As String = "update rplb_februari set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "06 Februari 2017" Then
                Call koneksi()
                Dim edit As String = "update rplb_februari set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "07 Februari 2017" Then
                Call koneksi()
                Dim edit As String = "update rplb_februari set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "08 Februari 2017" Then
                Call koneksi()
                Dim edit As String = "update rplb_februari set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "09 Februari 2017" Then
                Call koneksi()
                Dim edit As String = "update rplb_februari set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "10 Februari 2017" Then
                Call koneksi()
                Dim edit As String = "update rplb_februari set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "11 Februari 2017" Then
                Call koneksi()
                Dim edit As String = "update rplb_februari set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "12 Februari 2017" Then
                Call koneksi()
                Dim edit As String = "update rplb_februari set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "13 Februari 2017" Then
                Call koneksi()
                Dim edit As String = "update rplb_februari set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "14 Februari 2017" Then
                Call koneksi()
                Dim edit As String = "update rplb_februari set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "15 Februari 2017" Then
                Call koneksi()
                Dim edit As String = "update rplb_februari set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "16 Februari 2017" Then
                Call koneksi()
                Dim edit As String = "update rplb_februari set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "17 Februari 2017" Then
                Call koneksi()
                Dim edit As String = "update rplb_februari set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "18 Februari 2017" Then
                Call koneksi()
                Dim edit As String = "update rplb_februari set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "19 Februari 2017" Then
                Call koneksi()
                Dim edit As String = "update rplb_februari set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "20 Februari 2017" Then
                Call koneksi()
                Dim edit As String = "update rplb_februari set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "21 Februari 2017" Then
                Call koneksi()
                Dim edit As String = "update rplb_februari set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "22 Februari 2017" Then
                Call koneksi()
                Dim edit As String = "update rplb_februari set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "23 Februari 2017" Then
                Call koneksi()
                Dim edit As String = "update rplb_februari set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "24 Februari 2017" Then
                Call koneksi()
                Dim edit As String = "update rplb_februari set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "25 Februari 2017" Then
                Call koneksi()
                Dim edit As String = "update rplb_februari set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "26 Februari 2017" Then
                Call koneksi()
                Dim edit As String = "update rplb_februari set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "27 Februari 2017" Then
                Call koneksi()
                Dim edit As String = "update rplb_februari set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "28 Februari 2017" Then
                Call koneksi()
                Dim edit As String = "update rplb_februari set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()


            ElseIf Form1.DateTimePicker1.Text = "01 Maret 2017" Then
                Call koneksi()
                Dim edit As String = "update rplb_maret set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "02 Maret 2017" Then
                Call koneksi()
                Dim edit As String = "update rplb_maret set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "03 Maret 2017" Then
                Call koneksi()
                Dim edit As String = "update rplb_maret set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "04 Maret 2017" Then
                Call koneksi()
                Dim edit As String = "update rplb_maret set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "05 Maret 2017" Then
                Call koneksi()
                Dim edit As String = "update rplb_maret set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "06 Maret 2017" Then
                Call koneksi()
                Dim edit As String = "update rplb_maret set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "07 Maret 2017" Then
                Call koneksi()
                Dim edit As String = "update rplb_maret set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "08 Maret 2017" Then
                Call koneksi()
                Dim edit As String = "update rplb_maret set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "09 Maret 2017" Then
                Call koneksi()
                Dim edit As String = "update rplb_maret set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "10 Maret 2017" Then
                Call koneksi()
                Dim edit As String = "update rplb_maret set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "11 Maret 2017" Then
                Call koneksi()
                Dim edit As String = "update rplb_maret set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "12 Maret 2017" Then
                Call koneksi()
                Dim edit As String = "update rplb_maret set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "13 Maret 2017" Then
                Call koneksi()
                Dim edit As String = "update rplb_maret set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "14 Maret 2017" Then
                Call koneksi()
                Dim edit As String = "update rplb_maret set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "15 Maret 2017" Then
                Call koneksi()
                Dim edit As String = "update rplb_maret set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "16 Maret 2017" Then
                Call koneksi()
                Dim edit As String = "update rplb_maret set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "17 Maret 2017" Then
                Call koneksi()
                Dim edit As String = "update rplb_maret set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "18 Maret 2017" Then
                Call koneksi()
                Dim edit As String = "update rplb_maret set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "19 Maret 2017" Then
                Call koneksi()
                Dim edit As String = "update rplb_maret set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "20 Maret 2017" Then
                Call koneksi()
                Dim edit As String = "update rplb_maret set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "21 Maret 2017" Then
                Call koneksi()
                Dim edit As String = "update rplb_maret set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "22 Maret 2017" Then
                Call koneksi()
                Dim edit As String = "update rplb_maret set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "23 Maret 2017" Then
                Call koneksi()
                Dim edit As String = "update rplb_maret set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "24 Maret 2017" Then
                Call koneksi()
                Dim edit As String = "update rplb_maret set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "25 Maret 2017" Then
                Call koneksi()
                Dim edit As String = "update rplb_maret set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "26 Maret 2017" Then
                Call koneksi()
                Dim edit As String = "update rplb_maret set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "27 Maret 2017" Then
                Call koneksi()
                Dim edit As String = "update rplb_maret set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "28 Maret 2017" Then
                Call koneksi()
                Dim edit As String = "update rplb_maret set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "29 Maret 2017" Then
                Call koneksi()
                Dim edit As String = "update rplb_maret set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "30 Maret 2017" Then
                Call koneksi()
                Dim edit As String = "update rplb_maret set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "31 Maret 2017" Then
                Call koneksi()
                Dim edit As String = "update rplb_maret set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()


            ElseIf Form1.DateTimePicker1.Text = "01 April 2017" Then
                Call koneksi()
                Dim edit As String = "update rplb_april set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "02 April 2017" Then
                Call koneksi()
                Dim edit As String = "update rplb_april set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "03 April 2017" Then
                Call koneksi()
                Dim edit As String = "update rplb_april set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "04 April 2017" Then
                Call koneksi()
                Dim edit As String = "update rplb_april set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "05 April 2017" Then
                Call koneksi()
                Dim edit As String = "update rplb_april set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "06 April 2017" Then
                Call koneksi()
                Dim edit As String = "update rplb_april set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "07 April 2017" Then
                Call koneksi()
                Dim edit As String = "update rplb_april set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "08 April 2017" Then
                Call koneksi()
                Dim edit As String = "update rplb_april set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "09 April 2017" Then
                Call koneksi()
                Dim edit As String = "update rplb_april set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "10 April 2017" Then
                Call koneksi()
                Dim edit As String = "update rplb_april set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "11 April 2017" Then
                Call koneksi()
                Dim edit As String = "update rplb_april set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "12 April 2017" Then
                Call koneksi()
                Dim edit As String = "update rplb_april set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "13 April 2017" Then
                Call koneksi()
                Dim edit As String = "update rplb_april set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "14 April 2017" Then
                Call koneksi()
                Dim edit As String = "update rplb_april set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "15 April 2017" Then
                Call koneksi()
                Dim edit As String = "update rplb_april set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "16 April 2017" Then
                Call koneksi()
                Dim edit As String = "update rplb_april set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "17 April 2017" Then
                Call koneksi()
                Dim edit As String = "update rplb_april set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "18 April 2017" Then
                Call koneksi()
                Dim edit As String = "update rplb_april set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "19 April 2017" Then
                Call koneksi()
                Dim edit As String = "update rplb_april set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "20 April 2017" Then
                Call koneksi()
                Dim edit As String = "update rplb_april set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "21 April 2017" Then
                Call koneksi()
                Dim edit As String = "update rplb_april set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "22 April 2017" Then
                Call koneksi()
                Dim edit As String = "update rplb_april set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "23 April 2017" Then
                Call koneksi()
                Dim edit As String = "update rplb_april set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "24 April 2017" Then
                Call koneksi()
                Dim edit As String = "update rplb_april set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "25 April 2017" Then
                Call koneksi()
                Dim edit As String = "update rplb_april set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "26 April 2017" Then
                Call koneksi()
                Dim edit As String = "update rplb_april set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "27 April 2017" Then
                Call koneksi()
                Dim edit As String = "update rplb_april set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "28 April 2017" Then
                Call koneksi()
                Dim edit As String = "update rplb_april set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "29 April 2017" Then
                Call koneksi()
                Dim edit As String = "update rplb_april set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "30 April 2017" Then
                Call koneksi()
                Dim edit As String = "update rplb_april set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()


            ElseIf Form1.DateTimePicker1.Text = "01 Mei 2017" Then
                Call koneksi()
                Dim edit As String = "update rplb_mei set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "02 Mei 2017" Then
                Call koneksi()
                Dim edit As String = "update rplb_mei set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "03 Mei 2017" Then
                Call koneksi()
                Dim edit As String = "update rplb_mei set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "04 Mei 2017" Then
                Call koneksi()
                Dim edit As String = "update rplb_mei set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "05 Mei 2017" Then
                Call koneksi()
                Dim edit As String = "update rplb_mei set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "06 Mei 2017" Then
                Call koneksi()
                Dim edit As String = "update rplb_mei set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "07 Mei 2017" Then
                Call koneksi()
                Dim edit As String = "update rplb_mei set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "08 Mei 2017" Then
                Call koneksi()
                Dim edit As String = "update rplb_mei set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "09 Mei 2017" Then
                Call koneksi()
                Dim edit As String = "update rplb_mei set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "10 Mei 2017" Then
                Call koneksi()
                Dim edit As String = "update rplb_mei set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "11 Mei 2017" Then
                Call koneksi()
                Dim edit As String = "update rplb_mei set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "12 Mei 2017" Then
                Call koneksi()
                Dim edit As String = "update rplb_mei set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "13 Mei 2017" Then
                Call koneksi()
                Dim edit As String = "update rplb_mei set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "14 Mei 2017" Then
                Call koneksi()
                Dim edit As String = "update rplb_mei set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "15 Mei 2017" Then
                Call koneksi()
                Dim edit As String = "update rplb_mei set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "16 Mei 2017" Then
                Call koneksi()
                Dim edit As String = "update rplb_mei set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "17 Mei 2017" Then
                Call koneksi()
                Dim edit As String = "update rplb_mei set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "18 Mei 2017" Then
                Call koneksi()
                Dim edit As String = "update rplb_mei set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "19 Mei 2017" Then
                Call koneksi()
                Dim edit As String = "update rplb_mei set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "20 Mei 2017" Then
                Call koneksi()
                Dim edit As String = "update rplb_mei set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "21 Mei 2017" Then
                Call koneksi()
                Dim edit As String = "update rplb_mei set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "22 Mei 2017" Then
                Call koneksi()
                Dim edit As String = "update rplb_mei set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "23 Mei 2017" Then
                Call koneksi()
                Dim edit As String = "update rplb_mei set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "24 Mei 2017" Then
                Call koneksi()
                Dim edit As String = "update rplb_mei set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "25 Mei 2017" Then
                Call koneksi()
                Dim edit As String = "update rplb_mei set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "26 Mei 2017" Then
                Call koneksi()
                Dim edit As String = "update rplb_mei set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "27 Mei 2017" Then
                Call koneksi()
                Dim edit As String = "update rplb_mei set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "28 Mei 2017" Then
                Call koneksi()
                Dim edit As String = "update rplb_mei set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "29 Mei 2017" Then
                Call koneksi()
                Dim edit As String = "update rplb_mei set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "30 Mei 2017" Then
                Call koneksi()
                Dim edit As String = "update rplb_mei set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "31 Mei 2017" Then
                Call koneksi()
                Dim edit As String = "update rplb_mei set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()


            ElseIf Form1.DateTimePicker1.Text = "01 Juni 2017" Then
                Call koneksi()
                Dim edit As String = "update rplb_juni set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "02 Juni 2017" Then
                Call koneksi()
                Dim edit As String = "update rplb_juni set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "03 Juni 2017" Then
                Call koneksi()
                Dim edit As String = "update rplb_juni set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "04 Juni 2017" Then
                Call koneksi()
                Dim edit As String = "update rplb_juni set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "05 Juni 2017" Then
                Call koneksi()
                Dim edit As String = "update rplb_juni set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "06 Juni 2017" Then
                Call koneksi()
                Dim edit As String = "update rplb_juni set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "07 Juni 2017" Then
                Call koneksi()
                Dim edit As String = "update rplb_juni set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "08 Juni 2017" Then
                Call koneksi()
                Dim edit As String = "update rplb_juni set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "09 Juni 2017" Then
                Call koneksi()
                Dim edit As String = "update rplb_juni set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "10 Juni 2017" Then
                Call koneksi()
                Dim edit As String = "update rplb_juni set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "11 Juni 2017" Then
                Call koneksi()
                Dim edit As String = "update rplb_juni set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "12 Juni 2017" Then
                Call koneksi()
                Dim edit As String = "update rplb_juni set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "13 Juni 2017" Then
                Call koneksi()
                Dim edit As String = "update rplb_juni set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "14 Juni 2017" Then
                Call koneksi()
                Dim edit As String = "update rplb_juni set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "15 Juni 2017" Then
                Call koneksi()
                Dim edit As String = "update rplb_juni set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "16 Juni 2017" Then
                Call koneksi()
                Dim edit As String = "update rplb_juni set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "17 Juni 2017" Then
                Call koneksi()
                Dim edit As String = "update rplb_juni set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "18 Juni 2017" Then
                Call koneksi()
                Dim edit As String = "update rplb_juni set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "19 Juni 2017" Then
                Call koneksi()
                Dim edit As String = "update rplb_juni set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "20 Juni 2017" Then
                Call koneksi()
                Dim edit As String = "update rplb_juni set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "21 Juni 2017" Then
                Call koneksi()
                Dim edit As String = "update rplb_juni set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "22 Juni 2017" Then
                Call koneksi()
                Dim edit As String = "update rplb_juni set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "23 Juni 2017" Then
                Call koneksi()
                Dim edit As String = "update rplb_juni set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "24 Juni 2017" Then
                Call koneksi()
                Dim edit As String = "update rplb_juni set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "25 Juni 2017" Then
                Call koneksi()
                Dim edit As String = "update rplb_juni set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "26 Juni 2017" Then
                Call koneksi()
                Dim edit As String = "update rplb_juni set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "27 Juni 2017" Then
                Call koneksi()
                Dim edit As String = "update rplb_juni set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "28 Juni 2017" Then
                Call koneksi()
                Dim edit As String = "update rplb_juni set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "29 Juni 2017" Then
                Call koneksi()
                Dim edit As String = "update rplb_juni set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "30 Juni 2017" Then
                Call koneksi()
                Dim edit As String = "update rplb_juni set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            ElseIf Form1.DateTimePicker1.Text = "31 Juni 2017" Then
                Call koneksi()
                Dim edit As String = "update rplb_juni set Nama='" & TextBox2.Text & "' where NIS ='" & TextBox1.Text & "'"
                cmd = New OleDbCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call Form1.tampilgrid1()
            End If
        Catch ex As Exception
        End Try
        Me.Close()
    End Sub

    Private Sub TextBox1_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox1.KeyPress
        On Error Resume Next
        If e.KeyChar = Chr(13) Then
            Call koneksi()
            Call carikode2()
            If dr.HasRows Then
                Call ketemu3()

            End If
        End If
        If Not ((e.KeyChar >= "0" And e.KeyChar <= "9") Or e.KeyChar = vbBack) Then
            e.Handled = True
        End If
    End Sub

    Private Sub TextBox1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox1.TextChanged

    End Sub

    Private Sub edit2_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub
End Class