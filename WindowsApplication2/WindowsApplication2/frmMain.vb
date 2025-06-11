
Option Explicit On
Imports System.Data
Imports System.IO
Imports MySql.Data
Imports MySql.Data.MySqlClient
Imports System
Imports Excel = Microsoft.Office.Interop.Excel
Public Class frmMain
    Public mySchema As String = "zabbix"
    Private Sub Label2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label2.Click

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim connStr As String = "server=10.56.1.30;user=zabbix;database=zabbix;port=3306;password=Zabbix@2021;"
        Dim conn As MySqlConnection = New MySqlConnection(connStr)
        Dim strSQL As String
        Dim strUTimeBegin As Integer = 0
        Dim strUTimeEnd As Integer = 0
        Dim d1 As DateTime = DateTimePicker1.Value
        Dim d2 As DateTime = DateTimePicker2.Value

        'strUnixTimeBegin = DateTimePicker1.Value
        strUTimeBegin = (d1 - New DateTime(1970, 1, 1, 0, 0, 0)).TotalSeconds
        strUTimeEnd = (d2 - New DateTime(1970, 1, 1, 0, 0, 0)).TotalSeconds
        'strUTimeBegin = 1645507308
        ' strUTimeEnd = 1647926508
        '
        strSQL = "select hostmacro.value as group_name, hosts.hostid, hosts.host,hosts.name as host_name, items.name as item_name" & _
" , (" & strUTimeEnd & " - " & strUTimeBegin & ") as total_uptime " & _
" , (sum(if(events.value = 0,events.clock, 0))-sum(if(events.value = 1,events.clock, 0))) as downtime " & _
" , (((" & strUTimeEnd & " - " & strUTimeBegin & ") - (sum(if(events.value = 0,events.clock, 0))-sum(if(events.value = 1,events.clock, 0))))/(" & strUTimeEnd & " - " & strUTimeBegin & ") * 100 ) as avail " & _
"        from zabbix.events " & _
" left join zabbix.triggers on events.objectid = triggers.triggerid " & _
" left join zabbix.functions on triggers.triggerid = functions.triggerid " & _
" left join zabbix.items on functions.itemid = items.itemid " & _
" left join zabbix.hosts on items.hostid = hosts.hostid " & _
" left join zabbix.hosts_templates on hosts.hostid = hosts_templates.hostid " & _
" left join zabbix.hostmacro on hosts_templates.templateid = hostmacro.hostid " & _
" where items.key_ in ('icmpping','agent.ping','zabbix[host,snmp,available]') and hostmacro.macro='{$REPORT_GROUP}' " & _
" and events.clock > " & strUTimeBegin & " and events.clock < " & strUTimeEnd & "  " & _
" and not exists (select null from zabbix.trigger_tag where triggers.triggerid = trigger_tag.triggerid and trigger_tag.tag='report' and trigger_tag.value='no' ) " & _
" and events.eventid not in (select  eventid from events a, triggers b where a.objectid = b.triggerid and a.clock = b.lastchange and b.value=1) " & _
" group by hostmacro.value, hosts.hostid, hosts.host, hosts.name , items.name " & _
" order by hostmacro.value, hosts.name "
        '
        strSQL = "SELECT a.groupname, a.hostid, a.host, a.hostname , a.itemname " & _
            " , (" & strUTimeEnd & " - " & strUTimeBegin & ") as total_uptime " & _
" , (sum(if(value = 0,clock, 0))-sum(if(value = 1,clock, 0))) as downtime " & _
" , (((" & strUTimeEnd & " - " & strUTimeBegin & ") - (sum(if(value = 0,clock, 0))-sum(if(value = 1,clock, 0))))/(" & strUTimeEnd & " - " & strUTimeBegin & ") * 100 ) as avail " & _
            "FROM " & _
            " ( select hostmacro.value as groupname, hosts.hostid, hosts.host,hosts.name as hostname " & _
" , items.name as itemname , events.clock , events.value " & _
   "      from(zabbix.hosts) " & _
" left join zabbix.items on hosts.hostid = items.hostid " & _
" left join zabbix.functions on items.itemid = functions.itemid " & _
" left join zabbix.triggers on triggers.triggerid = functions.triggerid " & _
" left join events on events.objectid = triggers.triggerid  and events.clock > " & strUTimeBegin & " and events.clock < " & strUTimeEnd & " " & _
" left join zabbix.hosts_templates on hosts.hostid = hosts_templates.hostid  " & _
" left join zabbix.hostmacro on hosts_templates.templateid = hostmacro.hostid  " & _
" where items.key_ in ('icmpping','agent.ping','zabbix[host,snmp,available]') " & _
" and hostmacro.macro='{$REPORT_GROUP}'   " & _
" and not exists (select null from zabbix.trigger_tag where triggers.triggerid = trigger_tag.triggerid and trigger_tag.tag='report' and trigger_tag.value='no' ) " & _
" and events.eventid not in (select  eventid from events a, triggers b where a.objectid = b.triggerid and a.clock = b.lastchange and b.value=1)  " & _
 "        union all " &
" select hostmacro.value as groupname, hosts.hostid, hosts.host,hosts.name as hostname " & _
" , items.name as itemname , 0 as clock, 0 as value " & _
" from zabbix.hosts " & _
" left join zabbix.items on hosts.hostid = items.hostid " & _
" left join zabbix.functions on items.itemid = functions.itemid " & _
" left join zabbix.triggers on triggers.triggerid = functions.triggerid  " & _
" left join events on events.objectid = triggers.triggerid  and events.clock > " & strUTimeBegin & " and events.clock < " & strUTimeEnd & " " & _
" left join zabbix.hosts_templates on hosts.hostid = hosts_templates.hostid   " & _
" left join zabbix.hostmacro on hosts_templates.templateid = hostmacro.hostid   " & _
" where items.key_ in ('icmpping','agent.ping','zabbix[host,snmp,available]')  " & _
" and hostmacro.macro='{$REPORT_GROUP}'   " & _
" and not exists (select null from zabbix.trigger_tag where triggers.triggerid = trigger_tag.triggerid and trigger_tag.tag='report' and trigger_tag.value='no' )  " & _
" and events.eventid is null " & _
" ) a " & _
" group by a.groupname, a.hostid, a.host, a.hostname , a.itemname"

        '
        Try
            Console.WriteLine("Connecting to MySQL...")
            conn.Open()
            ' MsgBox(conn.State.ToString(), MsgBoxStyle.Information)

            Dim cmd As MySqlCommand = New MySqlCommand(strSQL, conn)
            Dim myReader As MySqlDataReader = cmd.ExecuteReader


            dg.Rows.Clear()
            dg.ColumnCount = 10
            'Format table
            dg.EnableHeadersVisualStyles = True
            dg.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize
            dg.RowHeadersVisible = False
            dg.Columns(0).Width = 100 'group_name
            dg.Columns(0).HeaderText = "Group name"
            dg.Columns(0).Visible = True
            dg.Columns(1).HeaderText = "HostID"
            dg.Columns(1).Width = 50
            dg.Columns(1).Visible = True
            dg.Columns(2).HeaderText = "Host"
            dg.Columns(2).Width = 150
            dg.Columns(2).Visible = False
            dg.Columns(3).HeaderText = "Host name"
            dg.Columns(3).Width = 200
            dg.Columns(3).Visible = True
            dg.Columns(4).HeaderText = "Key"
            dg.Columns(4).Width = 75
            dg.Columns(4).Visible = True
            dg.Columns(5).HeaderText = "Volume_up(s)"
            dg.Columns(5).Width = 75
            dg.Columns(5).Visible = True
            dg.Columns(6).HeaderText = "Volume_up"
            dg.Columns(6).Width = 75
            dg.Columns(6).Visible = True
            dg.Columns(7).HeaderText = "Downtime(s)"
            dg.Columns(7).Width = 75
            dg.Columns(7).Visible = True
            dg.Columns(8).HeaderText = "Downtime"
            dg.Columns(8).Width = 75
            dg.Columns(8).Visible = True
            dg.Columns(9).HeaderText = "Avail"
            dg.Columns(9).Width = 75
            dg.Columns(9).Visible = True
            dg.Width = 1000
            '
            Dim iDowntime As Long = 0
            Dim iHours As Integer = 0
            Dim iMinutes As Integer = 0
            Dim iSeconds As Integer = 0
            Dim strPreviousGroupname As String = ""
            '
            Dim iVolumeUp As Long = 0
            Dim idays2 As Integer = 0
            Dim iHours2 As Integer = 0
            Dim iSeconds2 As Integer = 0
            '
            Dim strAvail As String = ""
            '
            Do While myReader.Read = True
                '
                idays2 = 0
                iHours2 = 0
                iSeconds2 = 0
                'Neu co thay doi group, thi them vao dau dong
                If strPreviousGroupname <> myReader(0).ToString() Then
                    dg.Rows.Add(myReader(0))
                End If
                '
                iVolumeUp = myReader(5)
                iDowntime = myReader(6)
                strAvail = Format(myReader(7), "0.00")
                If iDowntime < 0 Then
                    'iVolumeUp = iDowntime
                    iDowntime = strUTimeEnd + iDowntime
                    strAvail = Format(((iVolumeUp - iDowntime) / iVolumeUp) * 100, "0.00")
                End If
                If iDowntime > 1495537095 Then
                    iDowntime = 0
                    strAvail = "N/A"
                End If
                iSeconds = iDowntime
                iHours = iSeconds / 3600
                iSeconds = iSeconds Mod 3600
                iMinutes = iSeconds / 60
                iSeconds = iSeconds Mod 60
                '
                iSeconds2 = iVolumeUp
                idays2 = iSeconds2 / (3600 * 24)
                iSeconds2 = iSeconds2 Mod (3600 * 24)
                iHours2 = iSeconds2 / 3600

                '
                dg.Rows.Add(myReader(0), myReader(1), myReader(2), myReader(3), myReader(4), myReader(5), idays2 & "days", iDowntime, iHours & "h" & iMinutes & "'" & iSeconds & """", strAvail)
                'Nho lai group cu
                strPreviousGroupname = myReader(0)
            Loop

        Catch ex As Exception
            MsgBox(ex.ToString(), vbCritical)
        End Try
        conn.Close()
        MsgBox("Đã thực hiện xong", MsgBoxStyle.Information)
    End Sub

    

    Private Sub DataGridView1_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs)

    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Me.Hide()
    End Sub

    Private Sub frmMain_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        DateTimePicker2.Value = DateTime.Now
        DateTimePicker1.Value = Now.Date.AddDays(-(Now.Day) + 1)

    End Sub

    Private Sub Button3_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Dim xlapp As Excel.Application
        Dim xlWorkBook As Excel.Workbook
        Dim xlWorkSheet As Excel.Worksheet
        Dim misValue As Object = System.Reflection.Missing.Value
        Dim i As Integer
        Dim j As Integer
        On Error Resume Next
        xlapp = New Excel.Application
        xlWorkBook = xlapp.Workbooks.Add(misValue)
        xlWorkSheet = CType(xlWorkBook.Sheets("Sheet1"), Excel.Worksheet)

        For k = 0 To dg.ColumnCount - 1
            xlWorkSheet.Cells(1, k + 1).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
            xlWorkSheet.Cells(1, k + 1) = dg.Columns(k).Name
        Next
        For i = 0 To dg.RowCount - 1
            For j = 0 To dg.ColumnCount - 1
                xlWorkSheet.Cells(i + 2, j + 1).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
               
                    xlWorkSheet.Cells(i + 2, j + 1) = dg(j, i).Value.ToString()

            Next
        Next

        Dim SaveFileDialog1 As New SaveFileDialog()
        SaveFileDialog1.Filter = "Execl files (*.xlsx)|*.xlsx"
        SaveFileDialog1.FilterIndex = 2
        SaveFileDialog1.RestoreDirectory = True
        If SaveFileDialog1.ShowDialog() = DialogResult.OK Then
            xlWorkSheet.SaveAs(SaveFileDialog1.FileName)
            MsgBox("Save file success")
        Else
            Return
        End If
        xlWorkBook.Close()
        xlapp.Quit()
    End Sub

    Private Sub Label3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label3.Click

    End Sub

    Private Sub Label4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label4.Click

    End Sub
End Class
