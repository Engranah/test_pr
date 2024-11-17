Imports System.Data.OleDb
Public Class frm_tanseeb_to
    Public Da As New SqlClient.SqlDataAdapter("select * from job order by job_nm ", Cn)
    Public Da1 As New SqlClient.SqlDataAdapter("select * from org order by org_nm", Cn)
    Public Da2 As New SqlClient.SqlDataAdapter("select * from org2 order by org_nm", Cn)
    Public Da3 As New SqlClient.SqlDataAdapter("select id,emp_nm from tanseeb_to order by emp_nm", Cn)
    Public Da4 As New SqlClient.SqlDataAdapter("select * from org order by org_nm", Cn)
    Dim flag As Boolean
    Dim cur_val As Boolean
    Dim sql As String
    Dim cmd As New SqlClient.SqlCommand
    Dim dr As SqlClient.SqlDataReader
    Dim drdt As SqlClient.SqlDataReader
    Dim drjob As SqlClient.SqlDataReader
    Dim drorg As SqlClient.SqlDataReader
    Dim drorg2 As SqlClient.SqlDataReader
    Dim drd As SqlClient.SqlDataReader
    Dim cur_no As Integer
    Dim ip As Integer
    Dim ds As New DataSet
    Dim ds2 As New DataSet


    Private Sub frm_tanseeb_to_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        If Cn.State <> ConnectionState.Open Then
            Cn.ConnectionString = Conn_Str
            Cn.Open()
        End If
        flag = True
        delet.Enabled = True
        Da.Fill(ds, "job")
        cmb_job.DataSource = ds
        cmb_job.ValueMember = ("job.job_no")
        cmb_job.DisplayMember = ("job.job_nm")
        cmb_job.SelectedIndex = -1

        Da1.Fill(ds, "org")
        cmb_org.DataSource = ds
        cmb_org.ValueMember = ("org.org_no")
        cmb_org.DisplayMember = ("org.org_nm")
        cmb_org.SelectedIndex = -1

        Da2.Fill(ds, "org2")
        cmb_org2.DataSource = ds
        cmb_org2.ValueMember = ("org2.org_no")
        cmb_org2.DisplayMember = ("org2.org_nm")
        cmb_org2.SelectedIndex = -1

        'If ds2.Tables.Count >= 1 Then
        '    ds2.Tables("fcors").Clear()
        'End If

        Da4.Fill(ds2, "org")
        Cmb_org3.DataSource = ds2
        Cmb_org3.ValueMember = ("org.org_no")
        Cmb_org3.DisplayMember = ("org.org_nm")
        Cmb_org3.SelectedIndex = -1



        Da3.Fill(ds, "tanseeb_to")
        cmb_nm.DataSource = ds
        cmb_nm.ValueMember = ("tanseeb_to.id")
        cmb_nm.DisplayMember = ("tanseeb_to.emp_nm")
        Panel1.Enabled = False
        Cmb_nm.SelectedIndex = -1
        Call header()
    End Sub
    Private Sub rr()
        cmb_job.SelectedIndex = -1
        cmb_org.SelectedIndex = -1
        cmb_nm.SelectedIndex = -1
        txt_nm.Text = ""
        cmb_org2.SelectedIndex = -1
        txt_no.Text = ""
        'cmb_note.SelectedIndex = -1
        dt1.Value = Now
        flag = True
    End Sub

    Private Sub add_Click(sender As Object, e As EventArgs) Handles add.Click
        Call rr()
        Call header2()
        cmb_note.SelectedIndex = -1
        flag = True
    End Sub

    Private Sub save_Click(sender As Object, e As EventArgs) Handles save.Click
        If Cn.State <> ConnectionState.Open Then
            Cn.ConnectionString = Conn_Str
            Cn.Open()
        End If


        Dim ran As String
        Dim cur_job, cur_org, cur_org2 As Integer
        'Dim cur_note As String
        If txt_nm.Text = "" Then
            MsgBox("ادخل اسم الموظف")
            Exit Sub
        Else
            If txt_no.Text = "" Then
                MsgBox("ادخل رقم كتاب التنسيب ")
                Exit Sub
            Else

                If cmb_org.Text = "" Then
                    MsgBox("ادخل دائرة المنسب منها ")
                    Exit Sub
                Else
                    If cmb_job.Text = "" Then
                        MsgBox("ادخل العنوان الوظيفي للموظف")
                        Exit Sub
                    Else
                        If cmb_org2.Text = "" Then
                            MsgBox("ادخل الدائرة المنسب اليها الموظف")
                            Exit Sub
                        Else

                            cur_job = cmb_job.SelectedValue
                            cur_org = cmb_org.SelectedValue
                            cur_org2 = cmb_org2.SelectedValue

                        End If
                    End If
                End If
            End If
            'If cmb_note.Text = "" Then
            '    cur_note = ""
            'Else
            '    cur_note = cmb_note.SelectedIndex
            'End If
            If flag = True Then

                'Dim drnm As SqlClient.SqlDataReader

r:              sql = "select * from tanseeb_to "
                cmd = New SqlClient.SqlCommand(sql, Cn)
                dr = cmd.ExecuteReader
                Do While dr.Read
                    cur_no = (dr!id)
                Loop
                cur_no = cur_no + 1


                sql = "INSERT INTO tanseeb_to (id,emp_nm,emp_job,emp_org_from,emp_org_to,book_no,book_dt,end_f) values (" & cur_no & ",'" + txt_nm.Text + "'," & cmb_job.SelectedValue & "," & cmb_org2.SelectedValue & "," & cmb_org.SelectedValue & ",'" + txt_no.Text + "','" + Format(dt1.Value, "yyyy/MM/dd") + "'," & 0 & ")"

                cmd = New SqlClient.SqlCommand(sql, Cn)
                cmd.ExecuteNonQuery()
                'Call header1()
                'Call detail()

                Call rr()


            Else


                ran = MsgBox("هل انت متأكد من الخزن", MsgBoxStyle.YesNo)
                If ran = vbYes Then
                    sql = " Update tanseeb_to set emp_nm= '" + txt_nm.Text + "',emp_job = " & cmb_job.SelectedValue & ",emp_org_to =" & cmb_org.SelectedValue & " ,emp_org_from =" & cmb_org2.SelectedValue & ", book_no = '" + txt_no.Text + "', book_dt = '" + Format(dt1.Value, "yyyy/MM/dd") + "' where id = " & ListView1.SelectedItems(0).Text & ""
                    cmd = New SqlClient.SqlCommand(sql, Cn)
                    cmd.ExecuteNonQuery()
                    Call header()
                    'Call detail()
                    Call rr()
                End If
            End If
        End If
        ds.Tables("tanseeb_to").Clear()

        Da3.Fill(ds, "tanseeb_to")
        cmb_nm.DataSource = ds
        cmb_nm.ValueMember = ("tanseeb_to.id")
        cmb_nm.DisplayMember = ("tanseeb_to.emp_nm")



        cmb_note.SelectedValue = -1

        cmb_nm.SelectedIndex = -1
        Call header2()
    End Sub
    Private Sub header()
        ListView1.Clear()
        Me.ListView1.Columns.Add("رقم الموظف", 0, HorizontalAlignment.Center)
        Me.ListView1.Columns.Add("اسم الموظف", 150, HorizontalAlignment.Left)
        Me.ListView1.Columns.Add("العنوان الوظيفي", 130, HorizontalAlignment.Left)
        Me.ListView1.Columns.Add("الجهة المنسب اليها", 150, HorizontalAlignment.Left)
        Me.ListView1.Columns.Add("الجهة المنسب منها", 150, HorizontalAlignment.Left)
        Me.ListView1.Columns.Add("رقم الكتاب", 100, HorizontalAlignment.Left)
        Me.ListView1.Columns.Add("تاريخ الكتاب", 80, HorizontalAlignment.Center)
      
    End Sub

    Private Sub delet_Click(sender As Object, e As EventArgs) Handles delet.Click
        Dim ran As String
        If flag = True Then
            Call rr()
        Else
            ran = MsgBox("هل انت متأكد من الحذف", MsgBoxStyle.YesNo)
            If ran = vbYes Then
                sql = "Delete from  tanseeb_to where id= " & ListView1.SelectedItems(0).Text & " "
                cmd = New SqlClient.SqlCommand(sql, Cn)
                cmd.ExecuteNonQuery()

                sql = "Delete from  notes_tanseeb_to where id= " & ListView1.SelectedItems(0).Text & " "
                cmd = New SqlClient.SqlCommand(sql, Cn)
                cmd.ExecuteNonQuery()

                Call rr()
                Call header2()
                Call header()


            Else
            End If
        End If
        ds.Tables("tanseeb_to").Clear()

        Da3.Fill(ds, "tanseeb_to")
        cmb_nm.DataSource = ds
        cmb_nm.ValueMember = ("tanseeb_to.id")
        cmb_nm.DisplayMember = ("tanseeb_to.emp_nm")


        cmb_nm.SelectedIndex = -1
        cmb_note.SelectedIndex = -1
    End Sub
    Private Sub detail()
        'Dim dro As SqlClient.SqlDataReader
        If Cn.State <> ConnectionState.Open Then
            Cn.ConnectionString = Conn_Str
            Cn.Open()
        End If

        Dim cur_job As String
        Dim cur_org As String
        Dim cur_org2 As String
        'Dim cur_sex As String



        'If cmb_nm.Text = "" Then : MsgBox("حقل البحث فارغة") : Exit Sub : End If


        sql = " SELECT id,emp_nm,emp_job,emp_org_from,emp_org_to,book_no,book_dt FROM tanseeb_to where emp_nm='" + cmb_nm.Text + "'"


        Try
            cmd = New SqlClient.SqlCommand(sql, Cn)
            drd = cmd.ExecuteReader
            ListView1.Items.Clear()


            Do While drd.Read
                Dim li As ListViewItem
                li = ListView1.Items.Add(drd!id)
                li.SubItems.Add(drd!emp_nm)
                If (drd!emp_job) <> 0 Then
                    Dim cmdjob As New SqlClient.SqlCommand("select job_nm from job where job_no =" & (drd!emp_job) & "", Cn)
                    cur_job = cmdjob.ExecuteScalar.ToString.Trim
                    li.SubItems.Add(cur_job)
                Else
                    li.SubItems.Add("")
                End If

                If (drd!emp_org_to) <> 0 Then
                    Dim cmdorg As New SqlClient.SqlCommand("select org_nm from org where org_no =" & (drd!emp_org_to) & "", Cn)
                    cur_org = cmdorg.ExecuteScalar.ToString.Trim
                    li.SubItems.Add(cur_org)
                Else
                    li.SubItems.Add("")
                End If

                If (drd!emp_org_from) <> 0 Then
                    Dim cmdorg2 As New SqlClient.SqlCommand("select org_nm from org2 where org_no =" & (drd!emp_org_from) & "", Cn)
                    cur_org2 = cmdorg2.ExecuteScalar.ToString.Trim
                    li.SubItems.Add(cur_org2)
                Else
                    li.SubItems.Add("")
                End If

                If (drd!book_no) <> "" Then
                    li.SubItems.Add(drd!book_no)
                Else
                    li.SubItems.Add("")
                End If

                If (drd!book_dt) <> "" Then
                    li.SubItems.Add(drd!book_dt)
                Else
                    li.SubItems.Add("")
                End If

                'If (drd!notes) <> "" Then

                '    cur_sex = (drd!notes)

                '    If cur_sex = "0" Then
                '        cur_sex = "تمديد تنسيب"
                '    ElseIf cur_sex = "1" Then
                '        cur_sex = "تغيير جهة تنسيب"
                '    End If
                '    li.SubItems.Add(cur_sex)

                'Else : li.SubItems.Add("")

                'End If





            Loop


        Catch e1 As SqlClient.SqlException

            ListView1.Text = e1.Message

        End Try

        drd = Nothing
    End Sub

    Private Sub ListView1_Click(sender As Object, e As EventArgs) Handles ListView1.Click
        If Cn.State <> ConnectionState.Open Then
            Cn.ConnectionString = Conn_Str
            Cn.Open()
        End If

        ToolStripButton2.Enabled = True
        ip = ListView1.SelectedItems(0).Text
        sql = "select id,emp_nm,emp_job,emp_org_from,emp_org_to,book_no,book_dt FROM tanseeb_to  where id= " & ListView1.SelectedItems(0).Text & " "
        cmd = New SqlClient.SqlCommand(sql, Cn)
        dr = cmd.ExecuteReader
        Do While dr.Read()
            txt_nm.Text = dr!emp_nm



            Dim cmdjob As New SqlClient.SqlCommand("select job_nm from job where job_no =" & dr!emp_job & "", Cn)
            drjob = cmdjob.ExecuteReader
            If drjob.Read Then
                cmb_job.SelectedValue = dr!emp_job
                cmb_job.Text = drjob(0)
            End If


            Dim cmdorg As New SqlClient.SqlCommand("select org_nm from org where org_no =" & dr!emp_org_to & "", Cn)
            drorg = cmdorg.ExecuteReader
            If drorg.Read Then
                cmb_org.SelectedValue = dr!emp_org_to
                cmb_org.Text = drorg(0)
            End If


            Dim cmdorg2 As New SqlClient.SqlCommand("select org_nm from org2 where org_no =" & dr!emp_org_from & "", Cn)
            drorg2 = cmdorg2.ExecuteReader
            If drorg2.Read Then
                cmb_org2.SelectedValue = dr!emp_org_from
                cmb_org2.Text = drorg2(0)
            End If



            txt_no.Text = dr!book_no
            dt1.Text = dr!book_dt

           
        Loop
        dr.Close()
        dr = Nothing
        flag = False
        Cmb_org3.SelectedValue = -1
        sql = "select * from notes_tanseeb_to where notes=" & 2 & " and id = " & ip & ""
        cmd = New SqlClient.SqlCommand(sql, Cn)
        dr = cmd.ExecuteReader
        If dr.Read Then
            MsgBox("التنسيب منتهي", MsgBoxStyle.OkOnly)
            ToolStripButton2.Enabled = False

            cmb_note.SelectedIndex = 2

            GroupBox1.Enabled = False
        End If

        GroupBox3.Enabled = True
        GroupBox3.Visible = True

        cmb_note.Enabled = False
        ToolStrip2.Enabled = True

    End Sub
    Private Sub closee_Click(sender As Object, e As EventArgs) Handles closee.Click
        Me.Dispose()
        Me.Close()
    End Sub

    Private Sub cmb_nm_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmb_nm.SelectedIndexChanged
        GroupBox1.Enabled = True
        cmb_note.SelectedIndex = -1
        Call header()
        Call header2()
        Call detail()
        txt_no2.Text = ""
        dt2.Value = Now
    End Sub

  

    Private Sub ToolStripButton1_Click(sender As Object, e As EventArgs) Handles ToolStripButton1.Click
        Panel1.Enabled = True
        cmb_note.Enabled = True
        txt_no2.Enabled = False
        dt2.Enabled = False
        Cmb_org3.Enabled = False
    End Sub

    Private Sub ToolStripButton2_Click(sender As Object, e As EventArgs) Handles ToolStripButton2.Click
        If Cn.State <> ConnectionState.Open Then
            Cn.ConnectionString = Conn_Str
            Cn.Open()
        End If
        Dim cur_note As Integer, ran, cur_org3 As String
        If cmb_note.Text = "" Then
            ran = MsgBox("حدد حالة التنسيب رجاءا", MsgBoxStyle.OkOnly)
        Else
            cur_note = cmb_note.SelectedIndex

        End If

        If Cmb_org3.SelectedIndex = -1 Then
            cur_org3 = ""
        Else : cur_org3 = Cmb_org3.SelectedValue
        End If

        If txt_no2.Text = "" Then
            ran = MsgBox("ادخل رقم الكتاب رجاءا", MsgBoxStyle.OkOnly)
        Else

        End If

        If cmb_note.SelectedIndex = 1 And Cmb_org3.SelectedIndex = -1 Then
            ran = MsgBox("ادخل الجهة المنسب اليها رجاءا", MsgBoxStyle.OkOnly)


        Else
            sql = "INSERT INTO notes_tanseeb_to (id,notes,book_no,book_dt,emp_org_to) values (" & ListView1.SelectedItems(0).Text & "," & cur_note & ",'" + txt_no2.Text + "','" + Format(dt2.Value, "yyyy/MM/dd") + "','" + cur_org3 + "')"
            cmd = New SqlClient.SqlCommand(sql, Cn)
            cmd.ExecuteNonQuery()

            If cur_note = 2 Then
                sql = " Update tanseeb_to set end_f= " & 1 & " where id = " & ListView1.SelectedItems(0).Text & ""
                cmd = New SqlClient.SqlCommand(sql, Cn)
                cmd.ExecuteNonQuery()
            End If
            Call header2()
            Call details()
            txt_no2.Text = ""
            dt2.Value = Now
            Cmb_org3.SelectedIndex = -1

        End If
    End Sub
    Private Sub header2()
        ListView2.Clear()
        Me.ListView2.Columns.Add("رمزالعنوان", 0, HorizontalAlignment.Left)
        Me.ListView2.Columns.Add("رقم الكتاب ", 100, HorizontalAlignment.Center)
        Me.ListView2.Columns.Add("تاريخ الكتاب", 100, HorizontalAlignment.Left)

    End Sub
    Private Sub details()
        If Cn.State <> ConnectionState.Open Then
            Cn.ConnectionString = Conn_Str
            Cn.Open()
        End If

        Dim cur_or As Integer = cmb_note.SelectedIndex


        sql = "select * from notes_tanseeb_to where notes=" & cur_or & " and id = " & ip & ""
        Try
            cmd = New SqlClient.SqlCommand(sql, Cn)
            drdt = cmd.ExecuteReader
            ListView2.Items.Clear()
            Do While drdt.Read
                Dim li As ListViewItem
                li = ListView2.Items.Add(drdt!id)
                li.SubItems.Add(drdt!book_no)
                li.SubItems.Add(drdt!book_dt)

            Loop

        Catch e1 As SqlClient.SqlException
            ListView2.Text = e1.Message
        End Try
        drdt.Close()
        drdt = Nothing
    End Sub

    Private Sub cmb_note_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmb_note.SelectedIndexChanged

        dt2.Enabled = True
        txt_no2.Enabled = True
        Panel1.Visible = True
        
        'تمديد تنسيب
        If cmb_note.SelectedIndex = 1 Then


            Label8.Visible = True
            Cmb_org3.Visible = True
            Cmb_org3.Enabled = True

            'تغيير جهة التنسيب
        ElseIf cmb_note.SelectedIndex = 2 Then
            Cmb_org3.Visible = False
            Label8.Visible = False


            'انهاء
        Else : cmb_note.SelectedIndex = 0
            Cmb_org3.Visible = False
            Label8.Visible = False
            ListView2.Visible = True


           
        End If
        Call header2()
        Call details()
    End Sub

  

    'Private Sub ListView2_Click(sender As Object, e As EventArgs) Handles ListView2.Click
    'If Cn.State <> ConnectionState.Open Then
    '    Cn.ConnectionString = Conn_Str
    '    Cn.Open()
    'End If
    'Dim cur_or As Integer = cmb_note.SelectedIndex
    'Dim ip1, ip2 As Integer
    'ip1 = Me.ListView2.SelectedItems(1).Text
    'ip2 = Me.ListView2.SelectedItems(2).Text
    'sql = "select * from notes_tanseeb_to where notes=" & cur_or & " and book_no= " & ListView2.Items(1).Text & " "
    'cmd = New SqlClient.SqlCommand(sql, Cn)
    'dr = cmd.ExecuteReader
    'Do While dr.Read
    '    txt_no2.Text = (dr!book_no)
    '    dt2.Text = (dr!book_dt)

    'Loop
    'flag = False
    'dr.Close()
    'dr = Nothing


    'End Sub

End Class