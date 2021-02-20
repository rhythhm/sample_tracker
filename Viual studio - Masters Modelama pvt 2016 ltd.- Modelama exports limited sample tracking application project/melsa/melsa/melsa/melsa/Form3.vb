Imports System
Imports System.Data
Imports System.Data.OleDb

Public Class Form3
    Dim buyer As String

    Dim d, r As Date
    Dim i, j As Integer
    Dim cn, cn1, cn2 As OleDbConnection
    Dim cmd, cmd1, cmd2 As OleDbCommand
    Dim dr As OleDbDataReader
    Dim da As OleDbDataAdapter
    Dim dt As DataTable
    Dim count, count1, count2 As Integer
    Dim sqlstr, constr, sqlstr1, sqlstr2, constr1 As String

    Public Sub cleartext()
        cbxxsf.Checked = False
        cbxsf.Checked = False
        cbsf.Checked = False
        cbmf.Checked = False
        cblf.Checked = False
        cbxlf.Checked = False
        cbxxlf.Checked = False

        cbxxspf.Checked = False
        cbxspf.Checked = False
        cbspf.Checked = False
        cbmpf.Checked = False
        cblpf.Checked = False
        cbxlpf.Checked = False
        cbxxlpf.Checked = False

        qxxsf.Text = "Qty"
        qxsf.Text = "Qty"
        qsf.Text = "Qty"
        qmf.Text = "Qty"
        qlf.Text = "Qty"
        qxlf.Text = "Qty"
        qxxlf.Text = "Qty"

        buyerlist.Text = "Select the Buyer for this Style"
        garmentlist.Text = "Select the Garment Type"
        iterationf.Text = "Select"
        iterationd.Text = "Select"
        txtstyle.Text = "Click to Enter Style Number"

    End Sub


    Private Sub RadioButton1_CheckedChanged(sender As Object, e As EventArgs) Handles rbmf.CheckedChanged
        cbxxsf.Visible = True
        cbxsf.Visible = True
        cbsf.Visible = True
        cbmf.Visible = True
        cblf.Visible = True
        cbxlf.Visible = True
        cbxxlf.Visible = True

        cbxxspf.Visible = False
        cbxspf.Visible = False
        cbspf.Visible = False
        cbmpf.Visible = False
        cblpf.Visible = False
        cbxlpf.Visible = False
        cbxxlpf.Visible = False
    End Sub

    Private Sub RadioButton2_CheckedChanged(sender As Object, e As EventArgs) Handles rbpf.CheckedChanged
        cbxxsf.Visible = False
        cbxsf.Visible = False
        cbsf.Visible = False
        cbmf.Visible = False
        cblf.Visible = False
        cbxlf.Visible = False
        cbxxlf.Visible = False

        cbxxspf.Visible = True
        cbxspf.Visible = True
        cbspf.Visible = True
        cbmpf.Visible = True
        cblpf.Visible = True
        cbxlpf.Visible = True
        cbxxlpf.Visible = True
    End Sub

    Private Sub btnclear_Click(sender As Object, e As EventArgs) Handles btnclear.Click
        cleartext()
    End Sub

    Private Sub buyerlist_SelectedIndexChanged(sender As Object, e As EventArgs) Handles buyerlist.SelectedIndexChanged
        buyer = buyerlist.Text
    End Sub
    Public Sub loadbuyer()
        Try
            constr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & Application.StartupPath & "\melsa1.accdb"
            sqlstr = "select name from buyerlist order by name"
            cn = New OleDbConnection(constr)

            cn.Open()
            cmd = New OleDbCommand(sqlstr, cn)
            dr = cmd.ExecuteReader
            While dr.Read
                buyerlist.Items.Add(dr("name"))
            End While
            cn.Close()

        Catch ex As Exception
            MsgBox("Please contact the admin" & ex.Message)
        End Try
    End Sub

    Public Sub loadgarment()
        Try
            constr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & Application.StartupPath & "\melsa1.accdb"
            sqlstr = "select name from garmentlist order by name"
            cn = New OleDbConnection(constr)

            cn.Open()
            cmd = New OleDbCommand(sqlstr, cn)
            dr = cmd.ExecuteReader
            While dr.Read
                garmentlist.Items.Add(dr("name"))
            End While
            cn.Close()

        Catch ex As Exception
            MsgBox("Please contact the admin" & ex.Message)
        End Try
    End Sub

    Private Sub Form3_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        loadbuyer()
        loadgarment()
    End Sub

    Private Sub cbxxsp_CheckedChanged(sender As Object, e As EventArgs) Handles cbxxspf.CheckedChanged
        If cbxxspf.Checked = True Then
            qxxsf.Enabled = True
        Else
            qxxsf.Enabled = False
        End If

    End Sub

    Private Sub qxs_SelectedIndexChanged(sender As Object, e As EventArgs) Handles qxsf.SelectedIndexChanged

    End Sub

    Private Sub cbxxs_CheckedChanged(sender As Object, e As EventArgs) Handles cbxxsf.CheckedChanged
        If cbxxsf.Checked = True Then
            qxxsf.Enabled = True
        Else
            qxxsf.Enabled = False
        End If
    End Sub

    Private Sub cbxsp_CheckedChanged(sender As Object, e As EventArgs) Handles cbxspf.CheckedChanged
        If cbxspf.Checked = True Then
            qxsf.Enabled = True
        Else
            qxsf.Enabled = False
        End If
    End Sub

    Private Sub cbxs_CheckedChanged(sender As Object, e As EventArgs) Handles cbxsf.CheckedChanged
        If cbxsf.Checked = True Then
            qxsf.Enabled = True
        Else
            qxsf.Enabled = False
        End If
    End Sub

    Private Sub cbsp_CheckedChanged(sender As Object, e As EventArgs) Handles cbspf.CheckedChanged
        If cbspf.Checked = True Then
            qsf.Enabled = True
        Else
            qsf.Enabled = False
        End If
    End Sub

    Private Sub cbs_CheckedChanged(sender As Object, e As EventArgs) Handles cbsf.CheckedChanged
        If cbsf.Checked = True Then
            qsf.Enabled = True
        Else
            qsf.Enabled = False
        End If
    End Sub

    Private Sub cbmp_CheckedChanged(sender As Object, e As EventArgs) Handles cbmpf.CheckedChanged
        If cbmpf.Checked = True Then
            qmf.Enabled = True
        Else
            qmf.Enabled = False
        End If
    End Sub

    Private Sub cbm_CheckedChanged(sender As Object, e As EventArgs) Handles cbmf.CheckedChanged
        If cbmf.Checked = True Then
            qmf.Enabled = True
        Else
            qmf.Enabled = False
        End If
    End Sub

    Private Sub cblp_CheckedChanged(sender As Object, e As EventArgs) Handles cblpf.CheckedChanged
        If cblpf.Checked = True Then
            qlf.Enabled = True
        Else
            qlf.Enabled = False
        End If
    End Sub

    Private Sub cbl_CheckedChanged(sender As Object, e As EventArgs) Handles cblf.CheckedChanged
        If cblf.Checked = True Then
            qlf.Enabled = True
        Else
            qlf.Enabled = False
        End If
    End Sub

    Private Sub cbxlp_CheckedChanged(sender As Object, e As EventArgs) Handles cbxlpf.CheckedChanged
        If cbxlpf.Checked = True Then
            qxlf.Enabled = True
        Else
            qxlf.Enabled = False
        End If
    End Sub

    Private Sub cbxl_CheckedChanged(sender As Object, e As EventArgs) Handles cbxlf.CheckedChanged
        If cbxlf.Checked = True Then
            qxlf.Enabled = True
        Else
            qxlf.Enabled = False
        End If
    End Sub

    Private Sub cbxxlp_CheckedChanged(sender As Object, e As EventArgs) Handles cbxxlpf.CheckedChanged
        If cbxxlpf.Checked = True Then
            qxxlf.Enabled = True
        Else
            qxxlf.Enabled = False
        End If
    End Sub

    Private Sub cbxxl_CheckedChanged(sender As Object, e As EventArgs) Handles cbxxlf.CheckedChanged
        If cbxxlf.Checked = True Then
            qxxlf.Enabled = True
        Else
            qxxlf.Enabled = False
        End If
    End Sub


    Private Sub btnhome_Click(sender As Object, e As EventArgs) Handles btnhome.Click
        Form2.Show()
        Me.Hide()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Try
            constr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & Application.StartupPath & "\melsa1.accdb"
            If cbfit.Checked = True Then
                If rbmf.Checked = True Then
                    sqlstr = "INSERT INTO fit (buyer_name,merchant_name,style_no,description,type_of_sample,iteration,xxs,xs,s,m,l,xl,xxl) VALUES ('" & buyerlist.Text & "','" & Form1.txtun.Text() & "','" & txtstyle.Text & "','" & garmentlist.Text & "','" & "FIT" & "','" & iterationf.Text & "','" & qxxsf.Text & "','" & qxsf.Text & "','" & qsf.Text & "','" & qmf.Text & "','" & qlf.Text & "','" & qxlf.Text & "','" & qxxlf.Text & "')"
                ElseIf rbpf.Checked = True Then
                    sqlstr = "INSERT INTO fit (buyer_name,merchant_name,style_no,description,type_of_sample,iteration,xxsp,xsp,sp,mp,lp,xlp,xxlp) VALUES ('" & buyerlist.Text & "','" & Form1.txtun.Text() & "','" & txtstyle.Text & "','" & garmentlist.Text & "','" & "FIT" & "','" & iterationf.Text & "','" & qxxsf.Text & "','" & qxsf.Text & "','" & qsf.Text & "','" & qmf.Text & "','" & qlf.Text & "','" & qxlf.Text & "','" & qxxlf.Text & "')"
                End If
            End If

            If cbdummy.Checked = True Then
                If rbmd.Checked = True Then
                    sqlstr1 = "INSERT INTO dummy (buyer_name,merchant_name,style_no,description,type_of_sample,iteration,xxs,xs,s,m,l,xl,xxl) VALUES ('" & buyerlist.Text & "','" & Form1.txtun.Text() & "','" & txtstyle.Text & "','" & garmentlist.Text & "','" & "DUMMY" & "','" & iterationd.Text & "','" & qxxsd.Text & "','" & qxsd.Text & "','" & qsd.Text & "','" & qmd.Text & "','" & qld.Text & "','" & qxld.Text & "','" & qxxld.Text & "')"
                ElseIf rbpd.Checked = True Then
                    sqlstr1 = "INSERT INTO dummy (buyer_name,merchant_name,style_no,description,type_of_sample,iteration,xxsp,xsp,sp,mp,lp,xlp,xxlp) VALUES ('" & buyerlist.Text & "','" & Form1.txtun.Text() & "','" & txtstyle.Text & "','" & garmentlist.Text & "','" & "DUMMY" & "','" & iterationd.Text & "','" & qxxsd.Text & "','" & qxsd.Text & "','" & qsd.Text & "','" & qmd.Text & "','" & qld.Text & "','" & qxld.Text & "','" & qxxld.Text & "')"
                End If
            End If
            cn = New OleDbConnection(constr)

            cn.Open()
            cmd = New OleDbCommand(sqlstr, cn)
            cmd1 = New OleDbCommand(sqlstr1, cn)
            count = cmd.ExecuteNonQuery()
            count1 = cmd1.ExecuteNonQuery()

            If count = 1 Then
                cleartext()
            End If
            If count1 = 1 Then
                cleartext()
            End If

            cn.Close()
        Catch ex As Exception
            MsgBox("Please contact to admin" & ex.Message)
        End Try
    End Sub

    Private Sub rbmd_CheckedChanged(sender As Object, e As EventArgs) Handles rbmd.CheckedChanged
        cbxxsd.Visible = True
        cbxsd.Visible = True
        cbsd.Visible = True
        cbmd.Visible = True
        cbld.Visible = True
        cbxld.Visible = True
        cbxxld.Visible = True

        cbxxspd.Visible = False
        cbxspd.Visible = False
        cbspd.Visible = False
        cbmpd.Visible = False
        cblpd.Visible = False
        cbxlpd.Visible = False
        cbxxlpd.Visible = False
    End Sub

    Private Sub rbpd_CheckedChanged(sender As Object, e As EventArgs) Handles rbpd.CheckedChanged
        cbxxsd.Visible = False
        cbxsd.Visible = False
        cbsd.Visible = False
        cbmd.Visible = False
        cbld.Visible = False
        cbxld.Visible = False
        cbxxld.Visible = False

        cbxxspd.Visible = True
        cbxspd.Visible = True
        cbspd.Visible = True
        cbmpd.Visible = True
        cblpd.Visible = True
        cbxlpd.Visible = True
        cbxxlpd.Visible = True
    End Sub

    Private Sub cbxxspd_CheckedChanged(sender As Object, e As EventArgs) Handles cbxxspd.CheckedChanged
        If cbxxspd.Checked = True Then
            qxxsd.Enabled = True
        Else
            qxxsd.Enabled = False
        End If
    End Sub

    Private Sub cbxspd_CheckedChanged(sender As Object, e As EventArgs) Handles cbxspd.CheckedChanged
        If cbxspd.Checked = True Then
            qxsd.Enabled = True
        Else
            qxsd.Enabled = False
        End If
    End Sub

    Private Sub cbspd_CheckedChanged(sender As Object, e As EventArgs) Handles cbspd.CheckedChanged
        If cbspd.Checked = True Then
            qsd.Enabled = True
        Else
            qsd.Enabled = False
        End If
    End Sub

    Private Sub cbmpd_CheckedChanged(sender As Object, e As EventArgs) Handles cbmpd.CheckedChanged
        If cbmpd.Checked = True Then
            qmd.Enabled = True
        Else
            qmd.Enabled = False
        End If
    End Sub

    Private Sub cblpd_CheckedChanged(sender As Object, e As EventArgs) Handles cblpd.CheckedChanged
        If cblpd.Checked = True Then
            qld.Enabled = True
        Else
            qld.Enabled = False
        End If
    End Sub

    Private Sub cbxlpd_CheckedChanged(sender As Object, e As EventArgs) Handles cbxlpd.CheckedChanged
        If cbxlpd.Checked = True Then
            qxld.Enabled = True
        Else
            qxld.Enabled = False
        End If
    End Sub

    Private Sub cbxxlpd_CheckedChanged(sender As Object, e As EventArgs) Handles cbxxlpd.CheckedChanged
        If cbxxlpd.Checked = True Then
            qxxld.Enabled = True
        Else
            qxxld.Enabled = False
        End If
    End Sub

    Private Sub cbxxsd_CheckedChanged(sender As Object, e As EventArgs) Handles cbxxsd.CheckedChanged
        If cbxxsd.Checked = True Then
            qxxsd.Enabled = True
        Else
            qxxsd.Enabled = False
        End If
    End Sub

    Private Sub cbxsd_CheckedChanged(sender As Object, e As EventArgs) Handles cbxsd.CheckedChanged
        If cbxsd.Checked = True Then
            qxsd.Enabled = True
        Else
            qxsd.Enabled = False
        End If
    End Sub

    Private Sub cbsd_CheckedChanged(sender As Object, e As EventArgs) Handles cbsd.CheckedChanged
        If cbsd.Checked = True Then
            qsd.Enabled = True
        Else
            qsd.Enabled = False
        End If
    End Sub

    Private Sub cbmd_CheckedChanged(sender As Object, e As EventArgs) Handles cbmd.CheckedChanged
        If cbmd.Checked = True Then
            qmd.Enabled = True
        Else
            qmd.Enabled = False
        End If
    End Sub

    Private Sub cbld_CheckedChanged(sender As Object, e As EventArgs) Handles cbld.CheckedChanged
        If cbld.Checked = True Then
            qld.Enabled = True
        Else
            qld.Enabled = False
        End If
    End Sub

    Private Sub cbxld_CheckedChanged(sender As Object, e As EventArgs) Handles cbxld.CheckedChanged
        If cbxld.Checked = True Then
            qxld.Enabled = True
        Else
            qxld.Enabled = False
        End If
    End Sub

    Private Sub cbxxld_CheckedChanged(sender As Object, e As EventArgs) Handles cbxxld.CheckedChanged
        If cbxxld.Checked = True Then
            qxxld.Enabled = True
        Else
            qxxld.Enabled = False
        End If
    End Sub

    Private Sub cbxlplr_CheckedChanged(sender As Object, e As EventArgs) Handles cbxlplr.CheckedChanged

    End Sub

   

    Private Sub CheckBox7_CheckedChanged(sender As Object, e As EventArgs) Handles cbxxspg.CheckedChanged

    End Sub

    Private Sub GroupBox1_Enter(sender As Object, e As EventArgs) Handles GroupBox1.Enter

    End Sub

    Private Sub qxsg_SelectedIndexChanged(sender As Object, e As EventArgs) Handles qxsg.SelectedIndexChanged

    End Sub

    Private Sub cbxspg_CheckedChanged(sender As Object, e As EventArgs) Handles cbxspg.CheckedChanged

    End Sub

    Private Sub cbspg_CheckedChanged(sender As Object, e As EventArgs) Handles cbspg.CheckedChanged

    End Sub

    Private Sub cbxlpg_CheckedChanged(sender As Object, e As EventArgs) Handles cbxlpg.CheckedChanged

    End Sub

    Private Sub cbxxlpg_CheckedChanged(sender As Object, e As EventArgs) Handles cbxxlpg.CheckedChanged

    End Sub

    Private Sub gbqr_Enter(sender As Object, e As EventArgs) Handles gbqr.Enter

    End Sub

    Private Sub cbxxspmr_CheckedChanged(sender As Object, e As EventArgs) Handles cbxxspmr.CheckedChanged

    End Sub

    Private Sub GPT_Enter(sender As Object, e As EventArgs) Handles GPT.Enter

    End Sub

    Private Sub CheckBox9_CheckedChanged(sender As Object, e As EventArgs) Handles cbspps.CheckedChanged

    End Sub

    Private Sub cbxspps_CheckedChanged(sender As Object, e As EventArgs) Handles cbxspps.CheckedChanged

    End Sub

    Private Sub cbsps_CheckedChanged(sender As Object, e As EventArgs) Handles cbsps.CheckedChanged

    End Sub

    Private Sub CheckBox6_CheckedChanged(sender As Object, e As EventArgs) Handles cbxsppp.CheckedChanged

    End Sub

    Private Sub CheckBox4_CheckedChanged(sender As Object, e As EventArgs) Handles cbmppp.CheckedChanged

    End Sub

    Private Sub CheckBox3_CheckedChanged(sender As Object, e As EventArgs) Handles cblppp.CheckedChanged

    End Sub

    Private Sub CheckBox2_CheckedChanged(sender As Object, e As EventArgs) Handles cbxlppp.CheckedChanged

    End Sub

    Private Sub cbxxsppp_CheckedChanged(sender As Object, e As EventArgs) Handles cbxxsppp.CheckedChanged

    End Sub

    Private Sub cbxxspp_CheckedChanged(sender As Object, e As EventArgs) Handles cbxxspp.CheckedChanged

    End Sub
End Class
