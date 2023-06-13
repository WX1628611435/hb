Imports System
Imports System.DateTime
Imports System.Data
Imports System.Data.SqlClient
Imports System.Configuration
Imports System.Net.Mail
Imports System.IO
Imports System.Net.NetworkInformation


Public Class Form1
    Dim add_D1000(1000), add_D101X(1000), add_D1021(1000), add_D14XX(1000), add_R2000(1000), add_R2001(1000), add_R2002(1000), add_R2003(1000), add_R5(1000), add_D1499(1000), add_D1599(1000) As Int32  ', add_D101X(1000)
    Dim sdb As SqlDB = New SqlDB
    Dim flag As Boolean
    Dim error_cont As Integer
    Dim m_id As String = My.Settings.machine_id
    Dim error_flag As Boolean

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Dim p As Ping = New Ping()
        Dim pr As PingReply
        pr = p.Send(My.Settings.PLC_ip)
        If (pr.Status <> IPStatus.Success) Then
            Save_Error("Plc_ip error: " & Now())
            Application.Exit()
        Else
            Dim p_date = Format(Date.Now.AddDays(-1), "yyyy-MM-dd")
            Dim p_date2 = Format(Date.Now.AddDays(-1), "yyyy/MM/dd")
            Dim p_date3 = Format(Date.Now.AddDays(-1), "yyyyMMdd")
            Save_Error("Load Start: " & Date.Now())

            Save_Error("Start: Me.PackageDataSet.PKG_ID" & Date.Now())
            'TODO: 這行程式碼會將資料載入 'PackageDataSet.PKG_ID' 資料表。您可以視需要進行移動或移除。
            Me.PKG_IDTableAdapter.Fill(Me.PackageDataSet.PKG_ID, My.Settings.machine_id)
            Save_Error("End: Me.PackageDataSet.PKG_ID" & Date.Now())

            'TODO: 這行程式碼會將資料載入 'Pud_dbDataSet.get_weight_today' 資料表。您可以視需要進行移動或移除。

            'TODO: 這行程式碼會將資料載入 'Ebi_db.pac_file' 資料表。您可以視需要進行移動或移除。
            '  Me.Pac_fileTableAdapter.Fill(Me.Ebi_db.pac_file)

            Save_Error("Start: Me.PackageDataSet.PERIOD_PACKING_DETAIL" & Date.Now())
            'TODO: 這行程式碼會將資料載入 'PackageDataSet.PERIOD_PACKING_DETAIL' 資料表。您可以視需要進行移動或移除。
            Me.PERIOD_PACKING_DETAILTableAdapter.Fill(Me.PackageDataSet.PERIOD_PACKING_DETAIL, My.Settings.machine_id, p_date)
            Save_Error("End: Me.PackageDataSet.PERIOD_PACKING_DETAIL" & Date.Now())

            Save_Error("Start: Me.PackageDataSet.PACKING_MACHINE" & Date.Now())
            'TODO: 這行程式碼會將資料載入 'PackageDataSet.PACKING_MACHINE' 資料表。您可以視需要進行移動或移除。
            Me.PACKING_MACHINETableAdapter.Fill(Me.PackageDataSet.PACKING_MACHINE)
            Save_Error("End: Me.PackageDataSet.PACKING_MACHINE" & Date.Now())

            Save_Error("Start: Me.PackageDataSet.PACKING_DETAIL" & Date.Now())
            'TODO: 這行程式碼會將資料載入 'PackageDataSet.PACKING_DETAIL' 資料表。您可以視需要進行移動或移除。
            Me.PACKING_DETAILTableAdapter.Fill(Me.PackageDataSet.PACKING_DETAIL, My.Settings.machine_id, p_date)
            Save_Error("End: Me.PackageDataSet.PACKING_DETAIL" & Date.Now())

            Save_Error("Start: Me.PackageDataSet.package_state" & Date.Now())
            'TODO: 這行程式碼會將資料載入 'PackageDataSet.package_state' 資料表。您可以視需要進行移動或移除。
            Me.Package_stateTableAdapter.Fill(Me.PackageDataSet.package_state, My.Settings.machine_id, p_date2)
            Save_Error("End: Me.PackageDataSet.package_state" & Date.Now())

            Save_Error("Start: Me.PackageDataSet.MaxSer" & Date.Now())
            'TODO: 這行程式碼會將資料載入 'PackageDataSet.MaxSer' 資料表。您可以視需要進行移動或移除。
            Me.MaxSerTableAdapter.Fill(Me.PackageDataSet.MaxSer, p_date3)
            Save_Error("End: Me.PackageDataSet.MaxSer" & Date.Now())

            Save_Error("Start: Me.PackageDataSet.ALERT_LOG" & Date.Now())
            'TODO: 這行程式碼會將資料載入 'PackageDataSet.ALERT_LOG' 資料表。您可以視需要進行移動或移除。
            Me.ALERT_LOGTableAdapter.Fill(Me.PackageDataSet.ALERT_LOG, My.Settings.machine_id, p_date)
            Save_Error("End: Me.PackageDataSet.ALERT_LOG" & Date.Now())

            If Timer2.Enabled = False Then
                Timer2.Enabled = True
            End If
            Save_Error("Load End: " & Date.Now())
        End If
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        AxActEasyIF5.Close()
        Timer2.Enabled = False
        Close()
    End Sub


    Sub ins_data()
        Dim com_msg As String = ""
        Dim p_date As String = ""

        com_msg = System.Convert.ToString(add_D1000(0), 2)
        com_msg = com_msg.PadLeft(16, "0")
        If Mid(com_msg, 1, 1) = "1" Then
            GaugeControl2.Enabled = True
            StateIndicatorComponent2.StateIndex = 2
        Else
            StateIndicatorComponent2.StateIndex = 3
        End If
        If Mid(com_msg, 2, 1) = "1" Then
            GaugeControl3.Enabled = True
            StateIndicatorComponent3.StateIndex = 1
        Else
            StateIndicatorComponent3.StateIndex = 3
        End If

        Array.Clear(add_R2000, 0, add_R2000.Length)
        If flag = True Then
            For a As Integer = 2000 To 2999 Step 10
                add_R2000(0) = 0
                add_R2001(0) = 0
                add_R2002(0) = 0
                add_R2003(0) = 0
                Dim b As Integer
                b = a Mod 10
                If b = 0 Then
                    Dim a0 As String = "R" & a.ToString
                    AxActEasyIF5.ReadDeviceBlock(a0, 1, add_R2000(0))
                    p_date = add_R2000(0)
                    p_date = p_date.PadLeft(4, "0")

                    If p_date = "0000" Then
                        mail_to(p_date)
                        Close()
                    Else
                        Dim year As String = Now.Year()
                        p_date = year & p_date
                    End If

                    Dim a1 As String = "R" & a + 1.ToString
                    AxActEasyIF5.ReadDeviceBlock(a1, 1, add_R2001(0))
                    Dim hhmm As String = add_R2001(0)
                    hhmm = hhmm.PadLeft(4, "0")
                    Dim hh As String = Mid(hhmm, 1, 2).ToString
                    Dim mm As String = Mid(hhmm, 3, 2).ToString

                    Dim a2 As String = "R" & a + 2.ToString
                    AxActEasyIF5.ReadDeviceBlock(a2, 1, add_R2002(0))
                    Dim ss As String = add_R2002(0)
                    ss = ss.PadLeft(2, "0")
                    Dim p_time As String = hh & ":" & mm & ":" & ss

                    Dim a3 As String = "R" & a + 3.ToString
                    AxActEasyIF5.ReadDeviceBlock(a3, 1, add_R2003(0))
                    Dim tot_ss As Integer = add_R2003(0)

                    If tot_ss <> 0 Then
                        Save_plc_detail("打包明細紀錄 " & "紀錄日期: " & Now() & " 日期位址:" & add_R2000(0) & " 打包日期: " & p_date & " 時分位址: " & a1 & "," & " 時分資料:" & add_R2001(0) & " 秒位址: " & a2 & " 秒資料：" & add_R2002(0) & " 歷時位址:" & a3 & " 歷時資料:" & add_R2003(0) & ",mid= " & m_id & ",p_date = " & p_date & ",p_time= " & p_time & ",tot_ss= " & tot_ss)
                        ins_prod_detail("", m_id, p_date, p_time, tot_ss)
                    End If
                End If
            Next

            day_packing_time(m_id)
            ins_err(m_id, p_date) '170428
        End If
    End Sub
    Sub mail_to(ByVal str)
        Dim err_str As String = str
        Try
            Dim mymail As New MailMessage
            mymail.From = New MailAddress("h0097@longchenpaper.com", "劉璟燁")
            'For Each to_addr As String In t_addr.Split(";")
            '    mymail.To.Add(to_addr)
            'Next
            'mymail.CC.Add("h0097@longchenpaper.com")
            mymail.SubjectEncoding = System.Text.Encoding.UTF8
            mymail.Subject = "日期錯誤：'" & err_str & "'"
            mymail.IsBodyHtml = True
            mymail.BodyEncoding = System.Text.Encoding.UTF8
            mymail.Body = "'" & err_str & "'"

            'Dim mysmtp As SmtpClient = New SmtpClient("www05.yuema.com.tw")
            Dim mysmtp As SmtpClient = New SmtpClient("192.168.4.9")
            'Dim smtpMail As New System.Net.Mail.SmtpClient()
            With mysmtp
                .Port = 25
                .EnableSsl = False
                .UseDefaultCredentials = False
                .DeliveryMethod = Net.Mail.SmtpDeliveryMethod.Network
                .Credentials = New System.Net.NetworkCredential("y006969", "y006969")
                '.Credentials = CredentialCache.DefaultNetworkCredentials("bpmflow", "bpmflow")
            End With
            mysmtp.Send(mymail)
        Catch ex As Exception
            'MsgBox(t_addr & "," & ex.Message)
        End Try
    End Sub
    Sub ins_state(ByVal m_id As String)
        Dim rec_sql As String = ""
        Dim e_str As String = ""
        Dim e_start As String = ""
        Dim ne_str As String = ""
        Dim ne_start As String = ""


        '本日紀錄
        '運轉時間
        Try
            Dim plc_rt As String = System.Convert.ToString(add_R5(0))
            plc_rt = plc_rt.PadLeft(4, "0")

            Dim tot_time As String = Mid(plc_rt, 1, 2) & "：" & Mid(plc_rt, 3, 4)

            Label1.Text = "運轉時間：" & Mid(plc_rt, 1, 2) & "時" & Mid(plc_rt, 3, 4) & "分"
            '打包數量
            Dim pack_num As String = "select sum(PACKAGE_NUM) from PERIOD_PACKING_DETAIL WHERE A_DATE = CONVERT(varchar(100), GETDATE(), 23) and machine_id = '" & m_id & "' "
            Dim tot_n As DataTable = sdb.GetDataTable(pack_num, CommandType.Text)
            Dim tot_num As String = tot_n.Rows(0)(0)
            Label5.Text = "打包數量：" & tot_num

            Dim old_sql As String = "select count(*) from package_state where machine_id = '" & My.Settings.machine_id & "'  and run_date = CONVERT(varchar(100), GETDATE(), 23)"
            Dim chk_dt As DataTable = sdb.GetDataTable(old_sql, CommandType.Text)
            Dim chk_ds As Integer = Val(chk_dt.Rows(0)(0) & "")
            If chk_ds > 0 Then
                rec_sql = "update package_state SET machine_state = '" & Label3.Text & "',run_time = '" & tot_time & "' ,package_num = '" & tot_num & "' where machine_id = '" & m_id & "'  and run_date = CONVERT(varchar(100), GETDATE(), 23) "
                sdb.RunCmd(rec_sql, CommandType.Text)
            Else
                rec_sql = "update package_state SET machine_state = '" & Label3.Text & "', run_time = '-' ,package_num = '0',run_date = CONVERT(varchar(100), GETDATE(), 23) where machine_id = '" & m_id & "'"
                sdb.RunCmd(rec_sql, CommandType.Text)
            End If

        Catch ex As Exception
            flag = False
            error_cont = error_cont + 1
            Save_Error("Ins state (Line " & error_cont & " Error) " & Mid(ex.Message, 1, ex.Message.Length - 1) & vbCrLf & "-->" & rec_sql)
        End Try


    End Sub

    Sub ins_err(ByVal m_id, ByVal p_date) '170428
        Dim d_error As String
        For i As Integer = 0 To 6
            If flag <> True Then
                Exit For
            End If
            add_D101X(0) = 0
            Dim ad_101X As String = "D101"
            Dim lb = i.ToString
            ad_101X = ad_101X & lb
            AxActEasyIF5.ReadDeviceBlock(ad_101X, 1, add_D101X(0))
            d_error = System.Convert.ToString(add_D101X(0), 2)
            error_log(ad_101X, m_id, d_error, p_date)
        Next

    End Sub
    Sub day_packing_time(ByVal m_id)
        Dim p_date_day As String = ""
        Dim pack_num_hour As String = ""
        Dim power_usage_hour As String = ""
        For k As Integer = 1 To 2
            If k = 1 Then
                AxActEasyIF5.ReadDeviceBlock("D1499", 1, add_D1499(0))
                p_date_day = add_D1499(0)
            Else
                AxActEasyIF5.ReadDeviceBlock("D1599", 1, add_D1599(0))
                p_date_day = add_D1599(0)
            End If

            p_date_day = p_date_day.PadLeft(4, "0")
            Dim year As String = Now.Year()
            p_date_day = year & p_date_day

            For i As Integer = 0 To 23
                add_D14XX(0) = 0
                If flag <> True Then
                    Exit For
                End If
                If k = 1 Then
                    pack_num_hour = "D14"
                    power_usage_hour = "D14"
                Else
                    pack_num_hour = "D15"
                    power_usage_hour = "D15"
                End If

                Dim la = i.ToString.PadLeft(2, "0")

                pack_num_hour = pack_num_hour & la

                AxActEasyIF5.ReadDeviceBlock(pack_num_hour, 1, add_D14XX(0))
                Dim pack_num As String = add_D14XX(0)

                Dim j = i + 50
                power_usage_hour = power_usage_hour & j.ToString
                AxActEasyIF5.ReadDeviceBlock(power_usage_hour, 1, add_D14XX(0))
                Dim power_usage As Integer = add_D14XX(0)

                Save_plc_period("各時段打包紀錄 " & "日期: " & Now() & "打包日期: " & p_date_day & " 位址: " & pack_num_hour & "," & " 打包數: " & pack_num & " 位址: " & power_usage_hour & "," & " 用電: " & power_usage)

                ins_period_packing_detail(m_id, p_date_day, i, pack_num, power_usage)
            Next

        Next
    End Sub

    Sub ins_period_packing_detail(ByVal m_id, ByVal p_date, ByVal i, ByVal pack_num, ByVal power_usage)
        Dim chk_cnt As DataTable = New DataTable
        Dim chk_ds As String = ""
        If flag = True Then
            Dim chk_str As String = "select count(*) from period_packing_detail "
            chk_str += " WHERE MACHINE_ID = '" & m_id & "' AND A_DATE = '" & p_date & "' AND PERIOD = '" & i & "' "
            Try
                chk_cnt = sdb.GetDataTable(chk_str, CommandType.Text)
                chk_ds = chk_cnt.Rows(0)(0)
            Catch ex As Exception
                flag = False
                error_cont = error_cont + 1
                Save_Error("Select (Line " & error_cont & " Error) " & Mid(ex.Message, 1, ex.Message.Length - 1) & vbCrLf & "-->" & chk_str)
            End Try


            If chk_ds = 0 Then
                Dim seqno As String = get_seqno("PERIOD_PACKING_DETAIL", p_date)

                Dim ins_period As String = "INSERT INTO PERIOD_PACKING_DETAIL (SEQ_NO,MACHINE_ID,A_DATE,PERIOD,PACKAGE_NUM"
                ins_period += ",POWER_USAGE,CREATION_DATE,CREATED_BY)"
                ins_period += " VALUES ('" & seqno & "','" & m_id & "','" & p_date & "','" & i & "','" & pack_num & "','" & power_usage & "' "
                ins_period += ",'" & p_date & "',-1)"
                Try
                    sdb.RunCmd(ins_period, CommandType.Text)
                Catch ex As Exception
                    flag = False
                    error_cont = error_cont + 1
                    Save_Error("Insert (Line " & error_cont & " Error) " & Mid(ex.Message, 1, ex.Message.Length - 1) & vbCrLf & "-->" & chk_str)
                End Try

            Else
                Dim ins_period As String = "UPDATE PERIOD_PACKING_DETAIL SET PACKAGE_NUM = '" & pack_num & "' , POWER_USAGE ='" & power_usage & "'"
                ins_period += " WHERE MACHINE_ID = '" & m_id & "' AND A_DATE = '" & p_date & "' AND PERIOD = '" & i & "' "
                Try
                    sdb.RunCmd(ins_period, CommandType.Text)
                Catch ex As Exception
                    flag = False
                    error_cont = error_cont + 1
                    Save_Error("Update (Line " & error_cont & " Error) " & Mid(ex.Message, 1, ex.Message.Length - 1) & vbCrLf & "-->" & ins_period)
                End Try

            End If
        End If
    End Sub

    Sub error_log(ByVal address, ByVal m_id, ByVal d_error, ByVal p_date) '170428
        d_error = d_error.PadLeft(16, "0")
        Dim err_msg As String
        Dim now_time As DateTime

        GaugeControl4.Enabled = False
        lab_err.Enabled = False
        Select Case address
            Case "D1010"
                For i = 1 To 16
                    If Mid(d_error, i, 1) = "1" Then

                        GaugeControl4.Enabled = True
                        lab_err.Enabled = True
                        StateIndicatorComponent4.StateIndex = 2
                        Select Case i
                            Case 1
                                err_msg = "NO2.M過載"
                                now_time = Format(Now(), "yyyy-MM-dd HH:mm:ss")
                                ins_err(m_id, err_msg, p_date, now_time)
                            Case 2
                                err_msg = "NO1.M過載"
                                now_time = Format(Now(), "yyyy-MM-dd HH:mm:ss")
                                ins_err(m_id, err_msg, p_date, now_time)
                            Case 3
                                err_msg = "油溫過高"
                                now_time = Format(Now(), "yyyy-MM-dd HH:mm:ss")
                                ins_err(m_id, err_msg, p_date, now_time)
                            Case 4
                                err_msg = "油位不足"
                                now_time = Format(Now(), "yyyy-MM-dd HH:mm:ss")
                                ins_err(m_id, err_msg, p_date, now_time)
                            Case 5
                                err_msg = "絞線部未閉合"
                                now_time = Format(Now(), "yyyy-MM-dd HH:mm:ss")
                                ins_err(m_id, err_msg, p_date, now_time)
                            Case 6
                                err_msg = "絞線門未關"
                                now_time = Format(Now(), "yyyy-MM-dd HH:mm:ss")
                                ins_err(m_id, err_msg, p_date, now_time)
                            Case 7
                                err_msg = "穿線右未關"
                                now_time = Format(Now(), "yyyy-MM-dd HH:mm:ss")
                                ins_err(m_id, err_msg, p_date, now_time)
                            Case 8
                                err_msg = "穿線左未關"
                                now_time = Format(Now(), "yyyy-MM-dd HH:mm:ss")
                                ins_err(m_id, err_msg, p_date, now_time)
                            Case 9
                                err_msg = "右護蓋未關"
                                now_time = Format(Now(), "yyyy-MM-dd HH:mm:ss")
                                ins_err(m_id, err_msg, p_date, now_time)
                            Case 10
                                err_msg = "左護蓋未關"
                                now_time = Format(Now(), "yyyy-MM-dd HH:mm:ss")
                                ins_err(m_id, err_msg, p_date, now_time)
                            Case 11
                                err_msg = "投料口未關_左"
                                now_time = Format(Now(), "yyyy-MM-dd HH:mm:ss")
                                ins_err(m_id, err_msg, p_date, now_time)
                            Case 12
                                err_msg = "紙料電眼異常"
                                now_time = Format(Now(), "yyyy-MM-dd HH:mm:ss")
                                ins_err(m_id, err_msg, p_date, now_time)
                            Case 13
                                err_msg = "塞紙電眼異常"
                                now_time = Format(Now(), "yyyy-MM-dd HH:mm:ss")
                                ins_err(m_id, err_msg, p_date, now_time)
                            Case 14
                                err_msg = "PLC電池異常"
                                now_time = Format(Now(), "yyyy-MM-dd HH:mm:ss")
                                ins_err(m_id, err_msg, p_date, now_time)
                            Case 15
                                err_msg = "按鈕卡住"
                                now_time = Format(Now(), "yyyy-MM-dd HH:mm:ss")
                                ins_err(m_id, err_msg, p_date, now_time)
                        End Select
                    Else
                        error_flag = False
                    End If
                Next
            Case "D1011"
                For i = 1 To 16
                    If Mid(d_error, i, 1) = 1 Then
                        lab_err.Enabled = True
                        GaugeControl4.Enabled = True
                        StateIndicatorComponent4.StateIndex = 2

                        Select Case i
                            Case 1
                                err_msg = "壓紙車未後定位"
                                now_time = Format(Now(), "yyyy-MM-dd HH:mm:ss")
                                ins_err(m_id, err_msg, p_date, now_time)
                            Case 2
                                err_msg = "散紙機未固定"
                                now_time = Format(Now(), "yyyy-MM-dd HH:mm:ss")
                                ins_err(m_id, err_msg, p_date, now_time)
                            Case 3
                                err_msg = "切線異常"
                                now_time = Format(Now(), "yyyy-MM-dd HH:mm:ss")
                                ins_err(m_id, err_msg, p_date, now_time)
                            Case 4
                                err_msg = "絞線過程異常"
                                now_time = Format(Now(), "yyyy-MM-dd HH:mm:ss")
                                ins_err(m_id, err_msg, p_date, now_time)
                            Case 5
                                err_msg = "穿線後過程異常"
                                now_time = Format(Now(), "yyyy-MM-dd HH:mm:ss")
                                ins_err(m_id, err_msg, p_date, now_time)
                            Case 6
                                err_msg = "穿線前過程異常"
                                now_time = Format(Now(), "yyyy-MM-dd HH:mm:ss")
                                ins_err(m_id, err_msg, p_date, now_time)
                            Case 7
                                err_msg = "壓紙車後過程異常"
                                now_time = Format(Now(), "yyyy-MM-dd HH:mm:ss")
                                ins_err(m_id, err_msg, p_date, now_time)
                            Case 8
                                err_msg = "壓紙車前過程異常"
                                now_time = Format(Now(), "yyyy-MM-dd HH:mm:ss")
                                ins_err(m_id, err_msg, p_date, now_time)
                            Case 9
                                err_msg = "NO2.散紙機過載"
                                now_time = Format(Now(), "yyyy-MM-dd HH:mm:ss")
                                ins_err(m_id, err_msg, p_date, now_time)
                            Case 10
                                err_msg = "NO1.散紙機過載"
                                now_time = Format(Now(), "yyyy-MM-dd HH:mm:ss")
                                ins_err(m_id, err_msg, p_date, now_time)
                            Case 11
                                err_msg = "散紙機移過載"
                                now_time = Format(Now(), "yyyy-MM-dd HH:mm:ss")
                                ins_err(m_id, err_msg, p_date, now_time)
                            Case 12
                                err_msg = "NO2.冷風過載"
                                now_time = Format(Now(), "yyyy-MM-dd HH:mm:ss")
                                ins_err(m_id, err_msg, p_date, now_time)
                            Case 13
                                err_msg = "NO1.冷風過載"
                                now_time = Format(Now(), "yyyy-MM-dd HH:mm:ss")
                                ins_err(m_id, err_msg, p_date, now_time)
                            Case 14
                                err_msg = "冷卻M回油過載"
                                now_time = Format(Now(), "yyyy-MM-dd HH:mm:ss")
                                ins_err(m_id, err_msg, p_date, now_time)
                            Case 15
                                err_msg = "絞線M過載"
                                now_time = Format(Now(), "yyyy-MM-dd HH:mm:ss")
                                ins_err(m_id, err_msg, p_date, now_time)
                            Case 16
                                err_msg = "NO3.M過載"
                                now_time = Format(Now(), "yyyy-MM-dd HH:mm:ss")
                                ins_err(m_id, err_msg, p_date, now_time)
                        End Select
                    Else
                        error_flag = False
                    End If
                Next
            Case "D1012"
                For i = 1 To 16
                    If Mid(d_error, i, 1) = 1 Then
                        GaugeControl4.Enabled = True
                        StateIndicatorComponent4.StateIndex = 2
                        lab_err.Enabled = True

                        Select Case i
                            Case 1
                                err_msg = "出料檢測_警告"
                                now_time = Format(Now(), "yyyy-MM-dd HH:mm:ss")
                                ins_err(m_id, err_msg, p_date, now_time)
                            Case 2
                                err_msg = "出料檢測_停車"
                                now_time = Format(Now(), "yyyy-MM-dd HH:mm:ss")
                                ins_err(m_id, err_msg, p_date, now_time)
                            Case 3
                                err_msg = "梱包無穿線完成"
                                now_time = Format(Now(), "yyyy-MM-dd HH:mm:ss")
                                ins_err(m_id, err_msg, p_date, now_time)
                            Case 4
                                err_msg = "梱包無絞線完成"
                                now_time = Format(Now(), "yyyy-MM-dd HH:mm:ss")
                                ins_err(m_id, err_msg, p_date, now_time)
                            Case 5
                                err_msg = "壓紙車無故收回"
                                now_time = Format(Now(), "yyyy-MM-dd HH:mm:ss")
                                ins_err(m_id, err_msg, p_date, now_time)
                            Case 6
                                err_msg = "穿線桿無故伸出"
                                now_time = Format(Now(), "yyyy-MM-dd HH:mm:ss")
                                ins_err(m_id, err_msg, p_date, now_time)
                            Case 7
                                err_msg = "穿線桿無故收回"
                                now_time = Format(Now(), "yyyy-MM-dd HH:mm:ss")
                                ins_err(m_id, err_msg, p_date, now_time)
                            Case 8
                                err_msg = "切線刀未在定位"
                                now_time = Format(Now(), "yyyy-MM-dd HH:mm:ss")
                                ins_err(m_id, err_msg, p_date, now_time)
                            Case 9
                                err_msg = "穿線位置無法辦識"
                                now_time = Format(Now(), "yyyy-MM-dd HH:mm:ss")
                                ins_err(m_id, err_msg, p_date, now_time)
                            Case 10
                                err_msg = "壓紙位置無法辦識"
                                now_time = Format(Now(), "yyyy-MM-dd HH:mm:ss")
                                ins_err(m_id, err_msg, p_date, now_time)
                            Case 11
                                err_msg = "變更操作模式"
                                now_time = Format(Now(), "yyyy-MM-dd HH:mm:ss")
                                ins_err(m_id, err_msg, p_date, now_time)
                            Case 12
                                err_msg = "散紙機未停無法出"
                                now_time = Format(Now(), "yyyy-MM-dd HH:mm:ss")
                                ins_err(m_id, err_msg, p_date, now_time)
                            Case 13
                                err_msg = "穿線未伸定位"
                                now_time = Format(Now(), "yyyy-MM-dd HH:mm:ss")
                                ins_err(m_id, err_msg, p_date, now_time)
                            Case 14
                                err_msg = "穿線未收定位"
                                now_time = Format(Now(), "yyyy-MM-dd HH:mm:ss")
                                ins_err(m_id, err_msg, p_date, now_time)
                            Case 15
                                err_msg = "絞線刀未定位"
                                now_time = Format(Now(), "yyyy-MM-dd HH:mm:ss")
                                ins_err(m_id, err_msg, p_date, now_time)
                            Case 16
                                err_msg = "壓紙車未前定位"
                                now_time = Format(Now(), "yyyy-MM-dd HH:mm:ss")
                                ins_err(m_id, err_msg, p_date, now_time)
                        End Select
                    Else
                        error_flag = False
                    End If
                Next
            Case "D1013"
                For i = 15 To 16
                    If Mid(d_error, i, 1) = 1 Then
                        GaugeControl4.Enabled = True
                        StateIndicatorComponent4.StateIndex = 2
                        lab_err.Enabled = True

                        Select Case i
                            Case 15
                                err_msg = "長度檢知無轉動"
                                now_time = Format(Now(), "yyyy-MM-dd HH:mm:ss")
                                ins_err(m_id, err_msg, p_date, now_time)
                            Case 16
                                err_msg = "遙控停止"
                                now_time = Format(Now(), "yyyy-MM-dd HH:mm:ss")
                                ins_err(m_id, err_msg, p_date, now_time)
                        End Select
                    End If
                Next
            Case "D1015"
                For i = 1 To 16
                    If Mid(d_error, i, 1) = 1 Then
                        GaugeControl4.Enabled = True
                        StateIndicatorComponent4.StateIndex = 2
                        lab_err.Enabled = True

                        Select Case i
                            Case 1
                                err_msg = "安全打包到達"
                                now_time = Format(Now(), "yyyy-MM-dd HH:mm:ss")
                                ins_err(m_id, err_msg, p_date, now_time)
                            Case 2
                                err_msg = "緊急停止_輸送帶"
                                now_time = Format(Now(), "yyyy-MM-dd HH:mm:ss")
                                ins_err(m_id, err_msg, p_date, now_time)
                            Case 3
                                err_msg = "緊急停止_絞線部"
                                now_time = Format(Now(), "yyyy-MM-dd HH:mm:ss")
                                ins_err(m_id, err_msg, p_date, now_time)
                            Case 4
                                err_msg = "緊急停止_其他"
                                now_time = Format(Now(), "yyyy-MM-dd HH:mm:ss")
                                ins_err(m_id, err_msg, p_date, now_time)
                            Case 5
                                err_msg = "緊急停止_操作箱"
                                now_time = Format(Now(), "yyyy-MM-dd HH:mm:ss")
                                ins_err(m_id, err_msg, p_date, now_time)
                            Case 6
                                err_msg = "緊急停止_主箱"
                                now_time = Format(Now(), "yyyy-MM-dd HH:mm:ss")
                                ins_err(m_id, err_msg, p_date, now_time)
                            Case 7
                                err_msg = "油溫過低"
                                now_time = Format(Now(), "yyyy-MM-dd HH:mm:ss")
                                ins_err(m_id, err_msg, p_date, now_time)
                            Case 8
                                err_msg = "緊急停止__過切"
                                now_time = Format(Now(), "yyyy-MM-dd HH:mm:ss")
                                ins_err(m_id, err_msg, p_date, now_time)
                            Case 9
                                err_msg = "緊急停止__斜板"
                                now_time = Format(Now(), "yyyy-MM-dd HH:mm:ss")
                                ins_err(m_id, err_msg, p_date, now_time)
                            Case 10
                                err_msg = "添加潤滑油"
                                now_time = Format(Now(), "yyyy-MM-dd HH:mm:ss")
                                ins_err(m_id, err_msg, p_date, now_time)
                            Case 11
                                err_msg = "清掃安全門"
                                now_time = Format(Now(), "yyyy-MM-dd HH:mm:ss")
                                ins_err(m_id, err_msg, p_date, now_time)
                            Case 12
                                err_msg = "2級保養時間到達"
                                now_time = Format(Now(), "yyyy-MM-dd HH:mm:ss")
                                ins_err(m_id, err_msg, p_date, now_time)
                            Case 13
                                err_msg = "1級保養時間到達"
                                now_time = Format(Now(), "yyyy-MM-dd HH:mm:ss")
                                ins_err(m_id, err_msg, p_date, now_time)
                            Case 14
                                err_msg = "液壓油更換時間到達"
                                now_time = Format(Now(), "yyyy-MM-dd HH:mm:ss")
                                ins_err(m_id, err_msg, p_date, now_time)
                            Case 15
                                err_msg = "真空檢知2_異常"
                                now_time = Format(Now(), "yyyy-MM-dd HH:mm:ss")
                                ins_err(m_id, err_msg, p_date, now_time)
                            Case 16
                                err_msg = "真空檢知1_異常"
                                now_time = Format(Now(), "yyyy-MM-dd HH:mm:ss")
                                ins_err(m_id, err_msg, p_date, now_time)
                        End Select
                    Else
                        error_flag = False
                    End If
                Next
            Case "D1016"
                For i = 1 To 16
                    If Mid(d_error, i, 1) = 1 Then
                        GaugeControl4.Enabled = True
                        StateIndicatorComponent4.StateIndex = 2
                        lab_err.Enabled = True
                        Select Case i
                            Case 1
                                err_msg = "安全打包到達"
                                now_time = Format(Now(), "yyyy-MM-dd HH:mm:ss")
                                ins_err(m_id, err_msg, p_date, now_time)
                            Case 2
                                err_msg = "緊急停止_絞線部"
                                now_time = Format(Now(), "yyyy-MM-dd HH:mm:ss")
                                ins_err(m_id, err_msg, p_date, now_time)
                            Case 3
                                err_msg = "壓力檢測異常"
                                now_time = Format(Now(), "yyyy-MM-dd HH:mm:ss")
                                ins_err(m_id, err_msg, p_date, now_time)
                            Case 4
                                err_msg = "緊急停止_輸送帶"
                                now_time = Format(Now(), "yyyy-MM-dd HH:mm:ss")
                                ins_err(m_id, err_msg, p_date, now_time)
                            Case 5
                                err_msg = "後退輔助閥_2"
                                now_time = Format(Now(), "yyyy-MM-dd HH:mm:ss")
                                ins_err(m_id, err_msg, p_date, now_time)
                            Case 6
                                err_msg = "警報器 異常"
                                now_time = Format(Now(), "yyyy-MM-dd HH:mm:ss")
                                ins_err(m_id, err_msg, p_date, now_time)
                            Case 7
                                err_msg = "警報器 啟動"
                                now_time = Format(Now(), "yyyy-MM-dd HH:mm:ss")
                                ins_err(m_id, err_msg, p_date, now_time)
                            Case 8
                                err_msg = "緊急停止__過切"
                                now_time = Format(Now(), "yyyy-MM-dd HH:mm:ss")
                                ins_err(m_id, err_msg, p_date, now_time)
                            Case 9
                                err_msg = "緊急停止__斜板"
                                now_time = Format(Now(), "yyyy-MM-dd HH:mm:ss")
                                ins_err(m_id, err_msg, p_date, now_time)
                            Case 10
                                err_msg = "全自動燈"
                                now_time = Format(Now(), "yyyy-MM-dd HH:mm:ss")
                                ins_err(m_id, err_msg, p_date, now_time)
                            Case 11
                                err_msg = "半自動燈"
                                now_time = Format(Now(), "yyyy-MM-dd HH:mm:ss")
                                ins_err(m_id, err_msg, p_date, now_time)
                            Case 12
                                err_msg = "手動燈"
                                now_time = Format(Now(), "yyyy-MM-dd HH:mm:ss")
                                ins_err(m_id, err_msg, p_date, now_time)
                            Case 13
                                err_msg = "捆包燈"
                                now_time = Format(Now(), "yyyy-MM-dd HH:mm:ss")
                                ins_err(m_id, err_msg, p_date, now_time)
                            Case 14
                                err_msg = "異常警報(BZ)"
                                now_time = Format(Now(), "yyyy-MM-dd HH:mm:ss")
                                ins_err(m_id, err_msg, p_date, now_time)
                            Case 15
                                err_msg = "後定位燈(G)"
                                now_time = Format(Now(), "yyyy-MM-dd HH:mm:ss")
                                ins_err(m_id, err_msg, p_date, now_time)
                            Case 16
                                err_msg = "捆綁燈(Y)"
                                now_time = Format(Now(), "yyyy-MM-dd HH:mm:ss")
                                ins_err(m_id, err_msg, p_date, now_time)
                        End Select
                    Else
                        error_flag = False
                    End If
                Next
        End Select
        If error_flag = False Then
            err_msg = "無異常"
            now_time = Format(Now(), "yyyy-MM-dd HH:mm:ss")
            ins_err(m_id, err_msg, p_date, now_time)
        End If
    End Sub
    Sub ins_err(ByVal m_id, ByVal err_msg, ByVal p_date, ByVal now_time)
        Dim e_update As String = ""
        Dim chk_cnt As DataTable = New DataTable
        Dim chk_ds As String = 0

        If flag = True Then
            Dim check_sql As String = "select count(*) from ALERT_LOG "
            check_sql += " WHERE MACHINE_ID = '" & m_id & "' AND A_DATE = '" & p_date & "' AND ALERT_TIME = '" & now_time & "'  "
            Try
                chk_cnt = sdb.GetDataTable(check_sql, CommandType.Text)
                chk_ds = chk_cnt.Rows(0)(0)
            Catch ex As Exception
                flag = False
                error_cont = error_cont + 1
                Save_Error("(Line " & error_cont & " Error) " & Mid(ex.Message, 1, ex.Message.Length - 1) & vbCrLf & "-->" & check_sql)
            End Try



            If chk_ds = 0 Then
                If err_msg <> "無異常" Then
                    Dim err_no As String = get_seqno("ALERT_LOG", p_date)
                    Dim err_str As String = "INSERT INTO [ALERT_LOG] ([SEQ_NO],[MACHINE_ID], [A_DATE],[ALERT_NAME],[ALERT_TIME],[CREATION_DATE],[CREATE_BY])"
                    err_str += " VALUES ('" & err_no & "','" & m_id & "','" & p_date & "','" & err_msg & "','" & now_time & "','" & p_date & "',-1)"
                    Try
                        sdb.RunCmd(err_str, CommandType.Text)
                    Catch ex As Exception
                        flag = False
                        error_cont = error_cont + 1
                        Save_Error("(Line " & error_cont & " Error) " & Mid(ex.Message, 1, ex.Message.Length - 1) & vbCrLf & "-->" & err_str)
                    End Try
                End If

                Dim chk_err_state As String = "select error_state,error_start from package_state"
                Dim chk_dt As DataTable = sdb.GetDataTable(chk_err_state, CommandType.Text)
                Dim chk_old_state As String = chk_dt.Rows(0)(0)
                Dim chk_old_start As String = chk_dt.Rows(0)(1)
                If chk_old_state <> err_msg And chk_old_start <> now_time.ToString Then
                    If err_msg = "無異常" Then
                        Try
                            e_update = "update package_state set error_state1 = error_state,error_start1 = error_start, "
                            e_update += " error_end1 =CONVERT(varchar(100), GETDATE(), 108), "
                            e_update += " error_state = '" & err_msg & "',error_start ='-',error_end = '-' "
                            e_update += "  where machine_id ='" & m_id & "' "
                            sdb.RunCmd(e_update, CommandType.Text)
                        Catch ex As Exception
                            flag = False
                            error_cont = error_cont + 1
                            Save_Error("(Line " & error_cont & " Error) " & Mid(ex.Message, 1, ex.Message.Length - 1) & vbCrLf & "-->" & e_update)
                        End Try
                    Else
                        Try
                            e_update = "update package_state set error_state1 = error_state,error_start1 = error_start, "
                            e_update += " error_end1 =CONVERT(varchar(100), GETDATE(), 108), "
                            e_update += " error_state = '" & err_msg & "',error_start = CONVERT(varchar(100),'" & now_time & "',108) "
                            e_update += "  where machine_id ='" & m_id & "' "
                            sdb.RunCmd(e_update, CommandType.Text)
                        Catch ex As Exception
                            flag = False
                            error_cont = error_cont + 1
                            Save_Error("(Line " & error_cont & " Error) " & Mid(ex.Message, 1, ex.Message.Length - 1) & vbCrLf & "-->" & e_update)
                        End Try
                    End If
                End If
            End If
        End If

    End Sub
    Function get_seqno(ByVal tbname, ByVal p_date) As String
        Dim dt As DataTable = New DataTable
        Dim arr As ArrayList = New ArrayList

        arr.Add(New SqlParameter("TBNAME", tbname))
        arr.Add(New SqlParameter("STRlen", 4))
        arr.Add(New SqlParameter("YMSTR", p_date))

        'Save_Error("Get_SeqNO:" & p_date & tbname)

        dt = sdb.GetDataTable("get_packageser", arr, CommandType.StoredProcedure)
        Dim pd_sn As String = dt.Rows(0)(0)
        Return pd_sn
    End Function

    Sub ins_prod_detail(ByVal seq_no As String, ByVal m_id As String, ByVal p_date As String, ByVal p_time As String, ByVal tot_ss As Integer)
        Dim seqno As String
        Dim INS_DE As String
        Dim chk_time As String = Mid(p_time, 1, 5)
        If flag = True Then
            Dim c_cnt As Integer
            Dim check_dt As DataTable = New DataTable
            Dim check_sql As String = "select count(*) from PACKING_DETAIL WHERE MACHINE_ID = '" & m_id & "' and P_DATE = convert(datetime, '" & p_date & "') and START_TIME like '%" & chk_time & "%' "
            Try
                check_dt = sdb.GetDataTable(check_sql, CommandType.Text)
                c_cnt = check_dt.Rows(0)(0)
            Catch ex As Exception
                flag = False
                error_cont = error_cont + 1
                Save_Error("c_cnt = " & c_cnt & "Select (Line " & error_cont & " Error) " & Mid(ex.Message, 1, ex.Message.Length - 1) & vbCrLf & "-->" & check_sql)
            End Try

            If c_cnt = 0 Then

                seqno = get_seqno("PACKING_DETAIL", p_date)



                INS_DE = "insert into PACKING_DETAIL (SEQ_NO,MACHINE_ID,P_DATE,START_TIME,ELAPSED_TIME,CREATION_DATE,CREATED_BY) "
                INS_DE += "VALUES ('" & seqno & "', '" & m_id & "',convert(datetime, '" & p_date & "'), '" & p_time & "','" & tot_ss & "', getdate(),-1)"
                Try
                    sdb.RunCmd(INS_DE, CommandType.Text)

                    Save_Error(INS_DE.ToString)

                Catch ex As Exception
                    flag = False
                    error_cont = error_cont + 1
                    Save_Error("Insert (Line " & error_cont & " Error) " & Mid(ex.Message, 1, ex.Message.Length - 1) & vbCrLf & "-->" & INS_DE)
                End Try
            End If

        End If
    End Sub

    Private Sub Save_plc_detail(ByVal ErrStr As String)
        Dim NewFile As Integer = FreeFile()
        FileOpen(NewFile, My.Application.Info.DirectoryPath & "\_plc_detail" & Format(Now(), "yyMMdd") & "_.log", OpenMode.Append)
        PrintLine(NewFile, ErrStr)
        FileClose(NewFile)
    End Sub

    Private Sub Save_plc_period(ByVal ErrStr As String)
        Dim NewFile As Integer = FreeFile()
        FileOpen(NewFile, My.Application.Info.DirectoryPath & "\_plc_period" & Format(Now(), "yyMMdd") & "_.log", OpenMode.Append)
        PrintLine(NewFile, ErrStr)
        FileClose(NewFile)
    End Sub
    Private Sub Save_plc_stats(ByVal ErrStr As String)
        Dim NewFile As Integer = FreeFile()
        FileOpen(NewFile, My.Application.Info.DirectoryPath & "\_plc_stats" & Format(Now(), "yyMMdd") & "_.log", OpenMode.Append)
        PrintLine(NewFile, ErrStr)
        FileClose(NewFile)
    End Sub

    Private Sub Save_Error(ByVal ErrStr As String)
        Dim NewFile As Integer = FreeFile()
        FileOpen(NewFile, My.Application.Info.DirectoryPath & "\_error_" & Format(Now(), "yyMMdd") & "_.log", OpenMode.Append)
        PrintLine(NewFile, ErrStr)
        FileClose(NewFile)
    End Sub
    Private Sub Timer2_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer2.Tick
        Timer2.Enabled = False

        Me.Text = m_id & "timer start" & Now()

        Save_Error("Timer2_Tick 1#" & Now())
        Me.Refresh()
        ins_pac_sql(m_id, Format(Now(), "yyyy/MM/dd"))
        Label3.Text = ""
        Save_Error("Timer2_Tick 1# end" & Now())

        'Dim auto_m As Integer = 0
        Dim auto_m As String = 0
        Dim a As Integer = 0
        Dim plc_date As Integer = 0
        error_flag = True
        Dim ti As Integer = 0

        Save_Error("Timer2_Tick 2# " & Now())

        While auto_m = 0 And ti < 3
            'While auto_m.Trim() = "" And ti < 3
            If AxActEasyIF5.Open() Then
                AxActEasyIF5.Close()
            End If
            AxActEasyIF5.Open()

            plc_date = Now.Day
            Dim plc_p As String = "R" & plc_date + 502
            Application.DoEvents()
            AxActEasyIF5.ActLogicalStationNumber = My.Settings.PLC_id
            AxActEasyIF5.ReadDeviceBlock("D1021", 15, add_D1021(0))
            auto_m = add_D1021(0)
            If auto_m > 0 Then
                'If auto_m.Trim() <> "" Then
                AxActEasyIF5.ReadDeviceBlock(plc_p, 1, add_R5(0))
            Else
                AxActEasyIF5.Close()
            End If
            ti = ti + 1
            Threading.Thread.Sleep(1000)
            Me.Text = ti & ":" & Now()
            Try
                Save_plc_stats("PLC 狀態:" & Me.Text)
            Catch ex As Exception
                Save_Error(Now() & " :plc狀態寫入錯誤!")
            End Try

            Me.Refresh()
        End While
        'Save_plc_stats("PLC 狀態:" & auto_m)

        Save_Error("Timer2_Tick 2# end " & Now())

        flag = True

        If Mid(auto_m.PadLeft(16, "0"), 13, 1) = 1 Then
            Label3.Text = "全自動"
        ElseIf Mid(auto_m.PadLeft(16, "0"), 14, 1) = 1 Then
            Label3.Text = "半自動"
        ElseIf Mid(auto_m.PadLeft(16, "0"), 15, 1) = 1 Then
            Label3.Text = "手動"
        Else
            Label3.Text = "關機中"
        End If

        'Select Case auto_m
        '    Case 28
        '        Label3.Text = "全自動"
        '    Case 4
        '        Label3.Text = "全自動"
        '    Case 12
        '        Label3.Text = "全自動"
        '    Case 20
        '        Label3.Text = "全自動"
        '    Case 25
        '        Label3.Text = "手動"
        '    Case 1
        '        Label3.Text = "手動"
        '    Case 9
        '        Label3.Text = "手動"
        '    Case 17
        '        Label3.Text = "手動"
        '    Case 26
        '        Label3.Text = "半自動"
        '    Case 2
        '        Label3.Text = "半自動"
        '    Case 10
        '        Label3.Text = "半自動"
        '    Case 18
        '        Label3.Text = "半自動"
        '    Case Else
        '        Label3.Text = "連線中斷"
        'End Select

        Try
            Save_plc_stats("連線:" & Label3.Text)
        Catch ex As Exception
            Save_Error(Now() & " :plc中文狀態寫入錯誤!")
        End Try

        If Label3.Text = "連線中斷" Then
            If AxActEasyIF5.Open() Then
                AxActEasyIF5.Close()
            End If
            Application.Exit()
            'End
        End If

        Save_Error("Timer2_Tick 3#" & Now())
        Label2.Text = Format(Now(), "yyyy-MM-dd HH:mm:ss") & " " & auto_m
        Save_plc_stats(Label2.Text)
        Save_Error("Timer2_Tick 3# end " & Now())

        Save_Error("Timer2_Tick 4# " & Now())
        Try
            Me.Text = m_id & "plc read start " & Now()
            Me.Refresh()
            ins_data()
            Save_plc_stats(Me.Text)
            Me.Text = m_id & "機台狀態寫入 " & Now()

            Me.Refresh()
            ins_state(m_id)
            Try
                Save_plc_stats(Me.Text)
            Catch ex As Exception
                Save_Error(Now() & " :機台狀態寫入錯誤!")
            End Try

            Me.Text = m_id & "plc read end " & Now()

            Me.Refresh()
            If AxActEasyIF5.Open() Then
                AxActEasyIF5.Close()
            End If
            Try
                Save_plc_stats(Me.Text)
            Catch ex As Exception
                Save_Error(Now() & " :plc 結束狀態寫入錯誤!")
            End Try
            Me.Text = m_id & "pac_file write start " & Now()

            Me.Refresh()
            Try
                Save_plc_stats(Me.Text)
            Catch ex As Exception
                Save_Error(Now() & " :pac_file 寫入狀態錯誤!")
            End Try
            ins_pac(m_id, Format(Now(), "yyyy/MM/dd"))
            'Close()
            'End
            Save_Error("Timer2_Tick 4# end " & Date.Now())
        Catch ex As Exception
            Save_Error("ex" & ex.ToString() & Date.Now())
        Finally
            Save_Error("Timer2_Tick 4# Finally end " & Date.Now())
            Application.Exit()
            'End
        End Try


        'Try
        'Catch ex As Exception
        '    If AxActEasyIF5.Open() Then
        '        AxActEasyIF5.Close()
        '    End If
        'End Try
        'Close()
        'End
    End Sub
    'Protected Sub shutdown()
    '    NotifyIcon1.Visible = False
    '    Application.Exit()
    'End Sub

    Private Sub Form1_Resize(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Resize
        If Me.WindowState = FormWindowState.Minimized Then
            Me.Hide()
            '  NotifyIcon1.Visible = True
            Me.Visible = False
        Else
            ' Me.ShowInTaskbar = True
        End If
    End Sub

    'Private Sub NotifyIcon1_DoubleClick(sender As System.Object, e As System.EventArgs) Handles NotifyIcon1.DoubleClick
    '    NotifyIcon1.Visible = False
    '    Me.Show()
    '    Me.WindowState = FormWindowState.Normal
    'End Sub

    '打包重量獨立新
    Sub ins_pac_sql(ByVal m_id, ByVal p_date)
        Dim pa_id As String = ""
        Dim w_ds As Double
        pa_id = get_pa_id(m_id)
        'Select Case m_id
        '    Case "PK001"
        '        pa_id = "B1"
        '    Case "PK002"
        '        pa_id = "C1"
        '    Case "PK003"
        '        pa_id = "I1"
        '    Case "PK004"
        '        pa_id = "J1"
        '    Case "PK007"
        '        pa_id = "E1"
        '    Case Else
        '        pa_id = "XX"
        'End Select

        Try
            'Dim cmd As New Odbc.OdbcCommand
            'cmd.CommandText = "set isolation to dirty read"
            Dim cnn As New Odbc.OdbcConnection
            cnn = Pac_fileTableAdapter.Connection
            'cmd.Connection = cnn
            'cnn.Open()
            'cmd.ExecuteNonQuery()

            Pac_fileTableAdapter.Fill(Ebi_db.pac_file, pa_id, CDate(p_date))
            Me.Get_weight_todayTableAdapter.Fill(Me.Pud_dbDataSet.get_weight_today)
            w_ds = Val(Pud_dbDataSet.get_weight_today.Rows(0)(0) & "")
            w_ds = Math.Round(w_ds / 1000, 2)
            Me.Text = "wight is " & w_ds.ToString
            Save_plc_stats(Me.Text)
            Me.Refresh()

            If w_ds <> 0 Then
                Try
                    If Pac_fileBindingSource.Count = 0 Then
                        Pac_fileBindingSource.AddNew()
                        Pac_fileBindingSource.Current("pac01") = pa_id
                        Pac_fileBindingSource.Current("pac02") = p_date
                        Pac_fileBindingSource.Current("pac10") = w_ds
                        Pac_fileBindingSource.Current("pacgrup") = "system"
                        'Pac_fileBindingSource.EndEdit()
                    Else
                        Pac_fileBindingSource.Current("pac10") = w_ds
                        ' Pac_fileTableAdapter.Update(Ebi_db.pac_file)
                    End If

                    Pac_fileBindingSource.EndEdit()
                    '      If Val(Pud_dbDataSet.get_weight_today.Rows(0)(0) & "") > 0 Then
                    Pac_fileTableAdapter.Update(Ebi_db.pac_file)
                    ' End If
                Catch ex As Exception

                End Try
                Pac_fileBindingSource.Filter = ""
            End If
        Catch ex As SqlException
            Save_Error(Now() & " :重量寫入pac_file 錯誤!" & pa_id & " " & p_date & " " & w_ds)
        End Try
    End Sub
    Function get_pa_id(ByVal m_id) As String
        Dim tmp_id As String = ""
        For i As Int16 = 0 To PKG_IDBindingSource.Count - 1
            If PKG_IDBindingSource.Item(i)("PLCID") = m_id Then
                tmp_id = PKG_IDBindingSource.Item(i)("ERPID")
                Exit For
            End If
        Next
        Return tmp_id
    End Function
    Sub ins_pac(ByVal m_id, ByVal p_date)
        Dim pa_id As String
        pa_id = get_pa_id(m_id)
        'Select Case m_id
        '    Case "PK001"
        '        pa_id = "B1"
        '    Case "PK002"
        '        pa_id = "C1"
        '    Case "PK003"
        '        pa_id = "I1"
        '    Case "PK004"
        '        pa_id = "J1"
        '    Case "PK007"
        '        pa_id = "E1"
        '    Case Else
        '        pa_id = "XX"
        'End Select

        Dim pakt As Double = 0  '打包時間(秒)
        Dim paknm As Integer = 0  '打包粒數(資料筆數)
        Dim paktt As Double = 0 'TT
        Dim power_ds As Double = 0 '用電量 kw/h
        Dim g_pakt As String = "select sum(elapsed_time) from packing_detail where p_date = '" & p_date & "' and MACHINE_ID = '" & m_id & "' "

        Dim g_paknm As String = "select count(*) from packing_detail where p_date = '" & p_date & "' and MACHINE_ID = '" & m_id & "'"

        Dim g_power As String = "select sum(power_usage) from period_packing_detail where a_date = '" & p_date & "' and MACHINE_ID = '" & m_id & "'"


        Try
            Dim pakt_dt As DataTable = sdb.GetDataTable(g_pakt, CommandType.Text)
            pakt = Val(pakt_dt.Rows(0)(0) & "")
            pakt = Math.Round(pakt / 3600, 2)

            Dim paknm_dt As DataTable = sdb.GetDataTable(g_paknm, CommandType.Text)
            paknm = Val(paknm_dt.Rows(0)(0) & "")

            Dim power_dt As DataTable = sdb.GetDataTable(g_power, CommandType.Text)
            power_ds = Val(power_dt.Rows(0)(0) & "")
            If paknm <> 0 Then
                paktt = Math.Round(pakt * 60 / paknm, 2)
                Try
                    If Pac_fileBindingSource.Count = 0 Then
                        Pac_fileBindingSource.AddNew()
                        Pac_fileBindingSource.Current("pac01") = pa_id
                        Pac_fileBindingSource.Current("pac02") = p_date
                        Pac_fileBindingSource.Current("pac03") = paknm
                        Pac_fileBindingSource.Current("pac04") = pakt
                        Pac_fileBindingSource.Current("pac05") = paktt
                        Pac_fileBindingSource.Current("pac06") = Label3.Text
                        Pac_fileBindingSource.Current("pac07") = ""
                        Pac_fileBindingSource.Current("pac08") = ""
                        Pac_fileBindingSource.Current("pac09") = power_ds
                        Pac_fileBindingSource.Current("pacgrup") = "system"

                    Else
                        Pac_fileBindingSource.Current("pac03") = paknm
                        Pac_fileBindingSource.Current("pac04") = pakt
                        Pac_fileBindingSource.Current("pac05") = paktt
                        Pac_fileBindingSource.Current("pac06") = Label3.Text
                        Pac_fileBindingSource.Current("pac07") = ""
                        Pac_fileBindingSource.Current("pac08") = ""
                        Pac_fileBindingSource.Current("pac09") = power_ds
                        Pac_fileBindingSource.Current("pacgrup") = "system"
                    End If
                    Pac_fileBindingSource.EndEdit()
                    Pac_fileTableAdapter.Update(Ebi_db.pac_file)
                Catch ex As Exception

                End Try
                Pac_fileBindingSource.Filter = ""
            End If
        Catch ex As SqlException
            Save_Error(Now() & " :pac_file 寫入錯誤!")
        End Try
    End Sub

    'Private Sub Button1_Click(sender As System.Object, e As System.EventArgs)
    '    Pac_fileTableAdapter.Fill(Ebi_db.pac_file, "B1", CDate("2015/11/05"))
    'End Sub

    'Private Sub Button3_Click_1(sender As System.Object, e As System.EventArgs)
    '    Pac_fileTableAdapter.Update(Ebi_db)
    'End Sub
End Class
