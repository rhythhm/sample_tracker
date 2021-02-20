Imports System
Imports System.Data
Imports System.Data.OleDb

Public Class Form1
    Dim cn As OleDbConnection
    Dim cmd As OleDbCommand
    Dim dr As OleDbDataReader
    Dim da As OleDbDataAdapter
    Dim dt As DataTable
    Dim count As Integer
    Dim sqlstr, constr As String

    Private Sub btnlogin_Click(sender As Object, e As EventArgs) Handles btnlogin.Click
        Try
            constr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & Application.StartupPath & "\melsa1.accdb"

            sqlstr = "select * from loginpage where  userid='" & txtun.Text & "' and password='" & txtpass.Text & "'and code ='" & "p" & "'"
            cn = New OleDbConnection(constr)
            cn.Open()
            cmd = New OleDbCommand(sqlstr, cn)
            dr = cmd.ExecuteReader
            If dr.Read Then
                Me.Visible = False
                Form2.Show()
            Else
                MsgBox("Invalid user name or password", vbCritical, "Search Message")
            End If
            cn.Close()
        Catch ex As Exception
            MsgBox("Please contact to admin" & ex.Message)
        End Try
    End Sub

    Private Sub btncancel_Click(sender As Object, e As EventArgs) Handles btncancel.Click
        Me.txtun.Clear()
        Me.txtpass.Clear()
    End Sub

    Private Sub txtpass_TextChanged(sender As Object, e As EventArgs) Handles txtpass.TextChanged

    End Sub
    Private Sub Label2_Click(sender As Object, e As EventArgs) Handles Label2.Click

    End Sub

    Private Sub txtun_TextChanged(sender As Object, e As EventArgs) Handles txtun.TextChanged

    End Sub
    Private Sub Label1_Click(sender As Object, e As EventArgs) Handles Label1.Click

    End Sub
    Private Sub Label3_Click(sender As Object, e As EventArgs) Handles Label3.Click

    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub
End Class
