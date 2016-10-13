Imports System.Data.SqlClient
Imports ExcelUtil
Imports Telerik.Web.UI
 
Public Class editable_grid
Inherits System.Web.UI.Page

    Dim radWin As New RadWindow

    Protected Sub Page_Init(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Init
        Dim oScriptMgr As RadScriptManager = Master.FindControl("RadScriptManager1")
        oScriptMgr.Services.Add(New ServiceReference("myService.asmx"))

        Dim meta As HtmlMeta = New HtmlMeta
        meta.HttpEquiv = "X-UA-Compatible"
        meta.Content = "IE=EmulateIE7"
        Me.Page.Header.Controls.AddAt(0, meta)

    End Sub

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        radWin.ID = "Window1"
        RadWindowManager1.Windows.Add(radWin)
        legendText.Text = Session("ReportName")
        'legendText2.Text = Session("ReportName")

        If Not Me.IsPostBack Then
            MKTSEG_Name.Attributes.Add("onchange", "MKSSelected(this, '" & MSValueHolder.ClientID & "', 'MKTSEG_Value');")
            DRA_Name.Attributes.Add("onchange", "MKSSelected(this, '" & DRAValueHolder.ClientID & "', 'DRA_Value');")
            dra_attr_name.Attributes.Add("onchange", "MKSSelected(this, '" & DRAAttrValueHolder.ClientID & "', 'Stub_DRA_Attr_Value');")
            Dim oRep As New RadGridState(RadGridReports)
            If Not IsNothing(Session("RepGrid")) Then
                oRep.LoadSettings((Session("RepGrid")))
            End If
            Dim oBan As New RadGridState(RadGridBanners)
            If Not IsNothing(Session("BanGrid")) Then
                oBan.LoadSettings((Session("BanGrid")))
            End If
            Dim oStub As New RadGridState(RadGridStubs)
            If Not IsNothing(Session("StubGrid")) Then
                oStub.LoadSettings((Session("StubGrid")))
            End If
        End If
        lblError.Text = ""

    End Sub

    Protected Sub RadTabStrip1_TabClick(ByVal sender As Object, ByVal e As Telerik.Web.UI.RadTabStripEventArgs) Handles RadTabStrip1.TabClick
        If e.Tab.Text = "Tables" Then
            RadGridReports.Rebind()
        End If
        If e.Tab.Text = "Banners" Then
            RadGridBanners.Rebind()
        End If
        If e.Tab.Text = "Stubs" Then
            RadGridStubs.Rebind()
        End If

        If e.Tab.Text = "New Table" Then
            Session("report_table_id") = 0
            LoadReportLookupsAndRecord()
        End If

        If e.Tab.Text = "New Banner" Then
            Session("banner_table_id") = 0
            LoadBannerLookupsAndRecord()
        End If

        If e.Tab.Text = "New Stub" Then
            Session("stub_table_id") = 0
            LoadStubLookupsAndRecord()
        End If

        If e.Tab.Text = "Re-sort Tables" Then
            LoadTree()
        End If

        If e.Tab.Text = "Import" Then
            Response.Redirect("editable_grid_excel_import.aspx")
        End If
    End Sub

    Protected Sub RadGridReports_NeedDataSource(ByVal source As Object, ByVal e As Telerik.Web.UI.GridNeedDataSourceEventArgs) Handles RadGridReports.NeedDataSource
        If RadTabStrip1.SelectedTab.Text = "Tables" Then
            Dim ds As DataSet = New DataSet
            Dim SQLConn As SqlConnection = GenUtils.DoConn("ParamRepConnString")

            Dim myCommand As New SqlCommand("frontend_TRR_ReportGrid", SQLConn)
            myCommand.CommandType = CommandType.StoredProcedure
            myCommand.CommandTimeout = 3600

            myCommand.Parameters.Add("@product_instance_id", SqlDbType.Int).Value = IIf(GenUtils.GetContextID("ProductIID").Length > 0, GenUtils.GetContextID("ProductIID"), 0)
            myCommand.Parameters.Add("@report_id", SqlDbType.Int).Value = IIf(GenUtils.GetContextID("ReportID").Length > 0, GenUtils.GetContextID("ReportID"), 0)

            Dim myData As New SqlDataAdapter(myCommand)
            myData.Fill(ds)

            myCommand.Dispose()
            GenUtils.UndoConn(SQLConn)

            RadGridReports.DataSource = ds.Tables(0)
        End If
    End Sub

    Protected Sub RadGridBanners_NeedDataSource(ByVal source As Object, ByVal e As Telerik.Web.UI.GridNeedDataSourceEventArgs) Handles RadGridBanners.NeedDataSource
        If RadTabStrip1.SelectedTab.Text = "Banners" Then
            Dim ds As DataSet = New DataSet
            Dim SQLConn As SqlConnection = GenUtils.DoConn("ParamRepConnString")

            Dim myCommand As New SqlCommand("frontend_TTR_BannerGrid", SQLConn)
            myCommand.CommandType = CommandType.StoredProcedure
            myCommand.CommandTimeout = 3600

            myCommand.Parameters.Add("@product_instance_id", SqlDbType.Int).Value = IIf(GenUtils.GetContextID("ProductIID").Length > 0, GenUtils.GetContextID("ProductIID"), 0)
            myCommand.Parameters.Add("@report_id", SqlDbType.Int).Value = IIf(GenUtils.GetContextID("ReportID").Length > 0, GenUtils.GetContextID("ReportID"), 0)

            Dim myData As New SqlDataAdapter(myCommand)
            myData.Fill(ds)

            myCommand.Dispose()
            GenUtils.UndoConn(SQLConn)

            RadGridBanners.DataSource = ds.Tables(0)
        End If
    End Sub

    Protected Sub RadGridStubs_NeedDataSource(ByVal source As Object, ByVal e As Telerik.Web.UI.GridNeedDataSourceEventArgs) Handles RadGridStubs.NeedDataSource
        If RadTabStrip1.SelectedTab.Text = "Stubs" Then
            Dim ds As DataSet = New DataSet
            Dim SQLConn As SqlConnection = GenUtils.DoConn("ParamRepConnString")

            Dim myCommand As New SqlCommand("frontend_TTR_StubGrid", SQLConn)
            myCommand.CommandType = CommandType.StoredProcedure
            myCommand.CommandTimeout = 3600

            myCommand.Parameters.Add("@report_id", SqlDbType.Int).Value = IIf(GenUtils.GetContextID("ReportID").Length > 0, GenUtils.GetContextID("ReportID"), 0)

            Dim myData As New SqlDataAdapter(myCommand)
            myData.Fill(ds)

            myCommand.Dispose()
            GenUtils.UndoConn(SQLConn)

            RadGridStubs.DataSource = ds.Tables(0)
        End If
    End Sub

    Protected Sub RadGridReports_Unload(ByVal sender As Object, ByVal e As System.EventArgs) Handles RadGridReports.Unload
        Dim oSet As New RadGridState(RadGridReports)
        Session("RepGrid") = oSet.SaveSettings()
    End Sub

    Protected Sub RadGridStubs_Unload(ByVal sender As Object, ByVal e As System.EventArgs) Handles RadGridStubs.Unload
        Dim oSet As New RadGridState(RadGridStubs)
        Session("StubGrid") = oSet.SaveSettings()
    End Sub

    Protected Sub RadGridBanners_Unload(ByVal sender As Object, ByVal e As System.EventArgs) Handles RadGridBanners.Unload
        Dim oSet As New RadGridState(RadGridBanners)
        Session("BanGrid") = oSet.SaveSettings()
    End Sub

    Protected Sub btnExport_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnExport.Click

        If chkBanners.Checked = False And chkReports.Checked = False And chkStubs.Checked = False Then
            lblError.Text = "* Please select at least one tab to export.<br />"
            Exit Sub
        End If

        Dim ds As DataSet = New DataSet
        Dim SQLConn As SqlConnection = GenUtils.DoConn("ParamRepConnString")

        Dim myCommand As New SqlCommand("frontend_TRR_GetExport", SQLConn)
        myCommand.CommandType = CommandType.StoredProcedure
        myCommand.CommandTimeout = 3600

        myCommand.Parameters.Add("@product_instance_id", SqlDbType.Int).Value = IIf(GenUtils.GetContextID("ProductIID").Length > 0, GenUtils.GetContextID("ProductIID"), 0)
        myCommand.Parameters.Add("@report_id", SqlDbType.Int).Value = IIf(GenUtils.GetContextID("ReportID").Length > 0, GenUtils.GetContextID("ReportID"), 0)
        myCommand.Parameters.Add("@getReports", SqlDbType.Bit).Value = chkReports.Checked
        myCommand.Parameters.Add("@getBanners", SqlDbType.Bit).Value = chkBanners.Checked
        myCommand.Parameters.Add("@getStubs", SqlDbType.Bit).Value = chkStubs.Checked

        Dim myData As New SqlDataAdapter(myCommand)
        myData.Fill(ds)

        myCommand.Dispose()
        GenUtils.UndoConn(SQLConn)

        ' -----------------------------------------------------------
        ' update table names and column names (for hidden columns)
        ' -----------------------------------------------------------
        If ds.Tables.Count > 0 Then
            For Each oTbl As DataTable In ds.Tables
                ' ------------------------------------------------------------------------------------
                ' add ACTION column to each table and put default value UPDATE
                '	this has to be done before altering the column names for hidden column indicators
                ' ------------------------------------------------------------------------------------
                If oTbl.Columns(0).Caption.ToLower = "table_id" Or oTbl.Columns(0).Caption.ToLower = "id" Or oTbl.Columns(0).Caption.ToLower = "report_stub_id" Then
                    Dim dc As DataColumn = New DataColumn("ACTION", System.Type.GetType("System.String"))
                    dc.DefaultValue = "UPDATE"
                    oTbl.Columns.Add(dc)
                    oTbl.Columns("ACTION").SetOrdinal(0)
                End If

                ' Rename Table name and mark columns for hidden fields (with "*" at the end of the column name)
                '	the column ordinal has moved up one (because of inserting ACTION column to the first), so those IDs are in 2nd column
                Select Case oTbl.Columns(1).Caption.ToLower
                    Case "table_id"
                        oTbl.TableName = "Tables"
                        ' add indicator (*) for hidden columns
                        oTbl.Columns("Table_ID").ColumnName += "*"
                        oTbl.Columns("Deliverable_ID").ColumnName += "*"
                        oTbl.Columns("Product_ID").ColumnName += "*"
                        oTbl.Columns("Product_Instance_Name").ColumnName += "*"
                    Case "id"
                        oTbl.TableName = "Banners"
                        ' add indicator (*) for hidden columns
                        oTbl.Columns("ID").ColumnName += "*"
                        oTbl.Columns("Deliverable_ID").ColumnName += "*"
                        oTbl.Columns("Product_ID").ColumnName += "*"
                        oTbl.Columns("Product_Instance_Name").ColumnName += "*"
                    Case "report_stub_id"
                        oTbl.TableName = "Stubs"
                        ' add indicator (*) for hidden columns
                        oTbl.Columns("Deliverable_ID").ColumnName += "*"
                        oTbl.Columns("Report_Stub_ID").ColumnName += "*"
                End Select

            Next oTbl

        End If

        ' ------------------------
        ' build Action Lookup
        ' ------------------------
        Dim orow As DataRow
        Dim dt As New DataTable
        dt.Columns.Add(New DataColumn("ACTION", GetType(System.String)))
        dt.TableName = "Lookup_Action"
        orow = dt.NewRow
        orow("ACTION") = "UPDATE"
        dt.Rows.Add(orow)
        orow = dt.NewRow
        orow("ACTION") = "DELETE"
        dt.Rows.Add(orow)
        orow = dt.NewRow
        orow("ACTION") = "INSERT"
        dt.Rows.Add(orow)

        ' --------------------------------
        ' add lookup table to the dataset
        ' --------------------------------
        ds.Tables.Add(dt)

        'Session("dsout") = ds
        'exportExcel(ds)

        Try
            Dim exportParam As ExportParameters = New ExportParameters With {
                .SqlDataSet = ds,
                .DownloadFileName = Session("ReportName").ToString & ".xlsx",
                .enableAlternateRowColor = False,
                .downloadExcelFile = False
            }
            ExcelLib.exportExcel(exportParam)
            Session("Excel_Export_Path") = GenUtils.getPhysicalPath(exportParam.serverDownloadFilePath)
            Context.Session("Excel_Export_Download_FileName") = exportParam.DownloadFileName
        Catch exception As ExcelLibException
        'lblError.Text = exception.Message
        End Try

        ' somehow the RadAjaxLoadingPanel won't close, redirect the download to a separate page
        radWin.NavigateUrl = "editable_grid_excel_export.aspx"
        radWin.Width = Unit.Pixel(25)
        radWin.Height = Unit.Pixel(25)
        radWin.Top = Unit.Pixel(1)
        radWin.Left = Unit.Pixel(1)
        radWin.VisibleStatusbar = False
        radWin.VisibleTitlebar = False
        radWin.Title = "Excel Export"
        radWin.Modal = False
        radWin.VisibleOnPageLoad = True

    End Sub

    Sub LoadTrendIDs()
        Dim ds As DataSet = New DataSet
        Dim SQLConn As SqlConnection = GenUtils.DoConn("SourceDBConnString")

        Dim myCommand As New SqlCommand("frontend_editable_grid_get_trend_ids", SQLConn)
        myCommand.CommandType = CommandType.StoredProcedure
        myCommand.CommandTimeout = 3600

        myCommand.Parameters.Add("@product_instance_id", SqlDbType.Int).Value = IIf(GenUtils.GetContextID("ProductIID").Length > 0, GenUtils.GetContextID("ProductIID"), 0)

        Dim myData As New SqlDataAdapter(myCommand)
        myData.Fill(ds)

        Source_System_Trend_ID.DataSource = ds.Tables(0)
        Source_System_Trend_ID.DataTextField = "trenddesc"
        Source_System_Trend_ID.DataValueField = "trend_id"
        Source_System_Trend_ID.DataBind()
        Source_System_Trend_ID.Items.Insert(0, New ListItem("Select One", ""))

        myCommand.Dispose()
        GenUtils.UndoConn(SQLConn)

    End Sub

    Sub LoadReportIDs()
        Dim ds As DataSet = New DataSet
        Dim SQLConn As SqlConnection = GenUtils.DoConn("ParamRepConnString")

        Dim myCommand As New SqlCommand("select distinct report_table_no from a_report_table where deliverable_id = @deliverable_id", SQLConn)
        myCommand.CommandType = CommandType.Text
        myCommand.CommandTimeout = 3600

        myCommand.Parameters.Add("@deliverable_id", SqlDbType.Int).Value = IIf(GenUtils.GetContextID("ReportID").Length > 0, GenUtils.GetContextID("ReportID"), 0)

        Dim myData As New SqlDataAdapter(myCommand)
        myData.Fill(ds)

        banner_report_table_no.DataSource = ds.Tables(0)
        banner_report_table_no.DataTextField = "report_table_no"
        banner_report_table_no.DataValueField = "report_table_no"
        banner_report_table_no.DataBind()
        ' banner_report_table_no.Items.Insert(0, New ListItem("Select One", ""))

        myCommand.Dispose()
        GenUtils.UndoConn(SQLConn)

    End Sub
    Sub LoadBannerLookupsAndRecord()
        btnBannerDelete.Enabled = True

        Dim ds As DataSet = New DataSet
        Dim SQLConn As SqlConnection = GenUtils.DoConn("SourceDBConnString")

        If IsNothing(Session("banner_table_id")) Then
            Session("banner_table_id") = 0
        End If

        If Session("banner_table_id") = 0 Then
            lblBannerID.Text = "0"
            lblBannerDeliverableID.Text = Session("ReportName")
            lblBannerProductID.Text = Session("ProductIID")
            lblBannerProductInstanceName.Text = Session("ProductIName")
        End If

        LoadTrendIDs()
        LoadReportIDs()

        Dim sAnswerVal As String = ""

        Dim myCommand As New SqlCommand("frontend_editable_grid_banner_lookups", SQLConn)
        myCommand.CommandType = CommandType.StoredProcedure
        myCommand.CommandTimeout = 3600

        myCommand.Parameters.Add("@domain_instance_id", SqlDbType.Int).Value = IIf(GenUtils.GetContextID("DomainIID").Length > 0, GenUtils.GetContextID("DomainIID"), 0)
        myCommand.Parameters.Add("@table_id", SqlDbType.Int).Value = Session("banner_table_id")

        Dim myData As New SqlDataAdapter(myCommand)
        myData.Fill(ds)

        If ds.Tables.Count > 0 Then

            Numerator.DataSource = ds.Tables(0)
            Numerator.DataTextField = "lookup_desc"
            Numerator.DataValueField = "lookup_value"
            Numerator.DataBind()
            Numerator.Items.Insert(0, New ListItem("Select One", ""))

            Denominator.DataSource = ds.Tables(1)
            Denominator.DataTextField = "lookup_desc"
            Denominator.DataValueField = "lookup_value"
            Denominator.DataBind()
            Denominator.Items.Insert(0, New ListItem("Select One", ""))

            Number_Format.DataSource = ds.Tables(2)
            Number_Format.DataTextField = "lookup_desc"
            Number_Format.DataValueField = "lookup_value"
            Number_Format.DataBind()
            Number_Format.Items.Insert(0, New ListItem("Select One", ""))

            Decimals.DataSource = ds.Tables(3)
            Decimals.DataTextField = "lookup_desc"
            Decimals.DataValueField = "lookup_value"
            Decimals.DataBind()
            Decimals.Items.Insert(0, New ListItem("Select One", ""))

            DRA_Name.DataSource = ds.Tables(4)
            DRA_Name.DataTextField = "DVName"
            DRA_Name.DataValueField = "derived_variable_name"
            DRA_Name.DataBind()
            DRA_Name.Items.Insert(0, New ListItem("Select One", ""))

            lstFormatType.DataSource = ds.Tables(5)
            lstFormatType.DataTextField = "lookup_desc"
            lstFormatType.DataValueField = "lookup_value"
            lstFormatType.DataBind()
            '   lstFormatType.Items.Insert(0, New ListItem("0 - None", "0"))

            If Session("banner_table_id") = 0 Then
                btnBannerDelete.Enabled = False

                Try
                    banner_report_table_no.SelectedIndex = 0
                Catch ex As Exception

                End Try

                txtBannerDisplay.Text = ""
                txtOrderBanner.Text = ""
                txtNumOfYears.Text = ""
                txtSuperHeader.Text = ""
                Try
                    Source_System_Trend_ID.SelectedIndex = 0
                Catch ex As Exception

                End Try

                Try
                    DRA_Name.SelectedIndex = 0
                Catch ex As Exception

                End Try
                Try

                Catch ex As Exception

                End Try
                'txtDRAValue.Text = ds.Tables(5).Rows(0).Item("DRA_Value").ToString

                txtIndicator.Text = ""
                txtObjectIndicator.Text = ""

                Try
                    Numerator.SelectedIndex = 0
                Catch ex As Exception

                End Try
                Try
                    Denominator.SelectedIndex = 0
                Catch ex As Exception

                End Try
                Try
                    Number_Format.SelectedIndex = 0
                Catch ex As Exception

                End Try
                Try
                    Decimals.SelectedIndex = 0
                Catch ex As Exception

                End Try

                txtControlColumn.Text = ""
                DRAValueHolder.InnerHtml = ""
            End If

            If ds.Tables.Count > 6 Then
                'This means we're loading an existing record too
                If ds.Tables(6).Rows.Count > 0 Then
                    sAnswerVal = ds.Tables(6).Rows(0).Item("dra_value").ToString
                End If
                If ds.Tables(6).Rows.Count > 0 And ds.Tables.Count > 5 Then
                    Dim sDrop As String = "<select id=""DRA_Value"" name=""DRA_Value"" style=""width:450px"">"
                    Dim s As New StringBuilder

                    For Each oRow As DataRow In ds.Tables(7).Rows
                        If oRow("answer_text").ToString = sAnswerVal Then
                            s.AppendLine("<option selected=""selected"" value=""" & sAnswerVal & """>" & sAnswerVal & "</option>")
                        Else
                            s.AppendLine("<option value=""" & oRow("answer_text").ToString & """>" & oRow("dvaname").ToString & "</option>")
                        End If
                    Next
                    s.AppendLine("</select>")
                    DRAValueHolder.InnerHtml = sDrop & s.ToString

                End If

                If ds.Tables(6).Rows.Count > 0 Then
                    Try
                        lblBannerID.Text = ds.Tables(6).Rows(0).Item("ID").ToString
                        lblBannerDeliverableID.Text = ds.Tables(6).Rows(0).Item("Deliverable_ID").ToString
                        lblBannerProductID.Text = ds.Tables(6).Rows(0).Item("Product_ID").ToString
                        lblBannerProductInstanceName.Text = ds.Tables(6).Rows(0).Item("Product_Instance_Name").ToString
                        Try
                            banner_report_table_no.SelectedValue = ds.Tables(6).Rows(0).Item("Report_Table_No").ToString
                        Catch ex As Exception

                        End Try

                        txtBannerDisplay.Text = ds.Tables(6).Rows(0).Item("Banner_Display").ToString
                        txtOrderBanner.Text = ds.Tables(6).Rows(0).Item("Order_Banner").ToString
                        txtNumOfYears.Text = ds.Tables(6).Rows(0).Item("No_of_Years").ToString
                        txtSuperHeader.Text = ds.Tables(6).Rows(0).Item("Super_Header").ToString
                        Try
                            lstFormatType.SelectedValue = ds.Tables(6).Rows(0).Item("format_type").ToString
                        Catch ex As Exception

                        End Try
                        Try
                            Source_System_Trend_ID.SelectedValue = ds.Tables(6).Rows(0).Item("Source_System_Trend_id").ToString
                        Catch ex As Exception

                        End Try

                        Try
                            DRA_Name.SelectedValue = ds.Tables(6).Rows(0).Item("dra_name").ToString
                        Catch ex As Exception

                        End Try
                        Try

                        Catch ex As Exception

                        End Try
                        'txtDRAValue.Text = ds.Tables(6).Rows(0).Item("DRA_Value").ToString

                        txtIndicator.Text = ds.Tables(6).Rows(0).Item("Indicator").ToString
                        txtObjectIndicator.Text = ds.Tables(6).Rows(0).Item("object_indicator").ToString()

                        Try
                            Numerator.SelectedValue = ds.Tables(6).Rows(0).Item("numerator").ToString
                        Catch ex As Exception

                        End Try
                        Try
                            Denominator.SelectedValue = ds.Tables(6).Rows(0).Item("denominator").ToString
                        Catch ex As Exception

                        End Try
                        Try
                            Number_Format.SelectedValue = ds.Tables(6).Rows(0).Item("number_format").ToString
                        Catch ex As Exception

                        End Try
                        Try
                            Decimals.SelectedValue = ds.Tables(6).Rows(0).Item("decimals").ToString
                        Catch ex As Exception

                        End Try

                        txtControlColumn.Text = ds.Tables(6).Rows(0).Item("Control_Column").ToString

                    Catch ex As Exception

                    End Try
                End If
            End If
        End If

        myCommand.Dispose()
        GenUtils.UndoConn(SQLConn)

    End Sub
    Sub LoadReportLookupsAndRecord()
        btnRDelete.Enabled = True

        Dim ds As DataSet = New DataSet
        Dim SQLConn As SqlConnection = GenUtils.DoConn("SourceDBConnString")
        If IsNothing(Session("report_table_id")) Then
            Session("report_table_id") = 0
        End If

        If Session("report_table_id") = 0 Then
            lblTableID.Text = "0"
            lblDeliverableID.Text = Session("ReportID")
            lblProductID.Text = Session("ProductIID")
            lblProductInstanceName.Text = Session("ProductIName")
        End If

        Dim sAnswerVal As String = ""

        Dim myCommand As New SqlCommand("frontend_editable_grid_report_lookups", SQLConn)
        myCommand.CommandType = CommandType.StoredProcedure
        myCommand.CommandTimeout = 3600

        myCommand.Parameters.Add("@domain_instance_id", SqlDbType.Int).Value = IIf(GenUtils.GetContextID("DomainIID").Length > 0, GenUtils.GetContextID("DomainIID"), 0)
        myCommand.Parameters.Add("@table_id", SqlDbType.Int).Value = Session("report_table_id")

        Dim myData As New SqlDataAdapter(myCommand)
        myData.Fill(ds)

        If ds.Tables.Count > 0 Then

            Table_Type_No.DataSource = ds.Tables(0)
            Table_Type_No.DataTextField = "lookup_desc"
            Table_Type_No.DataValueField = "lookup_value"
            Table_Type_No.DataBind()
            Table_Type_No.Items.Insert(0, New ListItem("Select One", ""))

            Table_Average_1.DataSource = ds.Tables(1)
            Table_Average_1.DataTextField = "lookup_desc"
            Table_Average_1.DataValueField = "lookup_value"
            Table_Average_1.DataBind()
            Table_Average_1.Items.Insert(0, New ListItem("Select One", ""))

            Table_Average_2.DataSource = ds.Tables(2)
            Table_Average_2.DataTextField = "lookup_desc"
            Table_Average_2.DataValueField = "lookup_value"
            Table_Average_2.DataBind()
            Table_Average_2.Items.Insert(0, New ListItem("Select One", ""))

            Table_Average_3.DataSource = ds.Tables(3)
            Table_Average_3.DataTextField = "lookup_desc"
            Table_Average_3.DataValueField = "lookup_value"
            Table_Average_3.DataBind()
            Table_Average_3.Items.Insert(0, New ListItem("Select One", ""))

            MKTSEG_Name.DataSource = ds.Tables(4)
            MKTSEG_Name.DataTextField = "DVName"
            MKTSEG_Name.DataValueField = "derived_variable_name"

            MKTSEG_Name.DataBind()
            MKTSEG_Name.Items.Insert(0, New ListItem("Select One", ""))

            If Session("report_table_id") = 0 Then
                btnRDelete.Enabled = False

                txtTopN.Text = ""
                txtReportTableNo.Text = ""
                txtReportName.Text = ""
                txtReportFooter1.Text = ""
                txtReportFooter2.Text = ""
                txtReportFooter3.Text = ""
                Try
                    MKTSEG_Name.SelectedIndex = 0
                Catch ex As Exception

                End Try

                txtMSDisplay.Text = ""
                txtRankType.Text = ""
                txtSortYear.Text = ""
                txtStubID.Text = ""
                Try
                    Table_Type_No.SelectedIndex = 0
                Catch ex As Exception

                End Try
                Try
                    Table_Average_1.SelectedIndex = 0
                Catch ex As Exception

                End Try
                Try
                    Table_Average_2.SelectedIndex = 0
                Catch ex As Exception

                End Try
                Try
                    Table_Average_3.SelectedIndex = 0
                Catch ex As Exception

                End Try

                txtTableBannerHeight.Text = ""
                MSValueHolder.InnerHtml = ""
                txtUniqueID.Text = ""
                txtGroupID.Text = ""
            End If

            If ds.Tables.Count > 5 Then
                'This means we're loading an existing record too
                If ds.Tables(5).Rows.Count > 0 Then
                    sAnswerVal = ds.Tables(5).Rows(0).Item("MKTSEG_value").ToString
                End If
                If ds.Tables(6).Rows.Count > 0 And ds.Tables.Count > 5 Then
                    Dim sDrop As String = "<select id=""MKTSEG_Value"" name=""MKTSEG_Value"" style=""width:450px"">"
                    Dim s As New StringBuilder

                    s.AppendLine("<option selected=""selected"" value=""" & sAnswerVal & """>" & sAnswerVal & "</option>")

                    For Each oRow As DataRow In ds.Tables(6).Rows
                        s.AppendLine("<option value=""" & oRow("answer_text").ToString & """>" & oRow("dvaname").ToString & "</option>")
                    Next
                    s.AppendLine("</select>")
                    MSValueHolder.InnerHtml = sDrop & s.ToString

                End If

                If ds.Tables(5).Rows.Count > 0 Then
                    Try
                        lblTableID.Text = ds.Tables(5).Rows(0).Item("table_id").ToString
                        lblDeliverableID.Text = ds.Tables(5).Rows(0).Item("deliverable_id").ToString
                        lblProductID.Text = ds.Tables(5).Rows(0).Item("product_id").ToString
                        lblProductInstanceName.Text = ds.Tables(5).Rows(0).Item("Product_Instance_Name").ToString

                        txtTopN.Text = ds.Tables(5).Rows(0).Item("topn").ToString
                        txtReportTableNo.Text = ds.Tables(5).Rows(0).Item("report_table_no").ToString
                        txtReportName.Text = ds.Tables(5).Rows(0).Item("report_name").ToString
                        txtReportFooter1.Text = ds.Tables(5).Rows(0).Item("report_footer").ToString
                        txtReportFooter2.Text = ds.Tables(5).Rows(0).Item("report_footer2").ToString
                        txtReportFooter3.Text = ds.Tables(5).Rows(0).Item("report_footer3").ToString
                        Try
                            MKTSEG_Name.SelectedValue = ds.Tables(5).Rows(0).Item("MKTSEG_name").ToString
                        Catch ex As Exception

                        End Try
                        'Try
                        '    MKTSEG_Value.SelectedValue = ds.Tables(5).Rows(0).Item("MKTSEG_value").ToString
                        'Catch ex As Exception

                        'End Try

                        txtMSDisplay.Text = ds.Tables(5).Rows(0).Item("MKTSEG_display").ToString
                        txtRankType.Text = ds.Tables(5).Rows(0).Item("rank_type").ToString
                        txtSortYear.Text = ds.Tables(5).Rows(0).Item("sort_year").ToString
                        txtStubID.Text = ds.Tables(5).Rows(0).Item("stub_id").ToString
                        Try
                            Table_Type_No.SelectedValue = ds.Tables(5).Rows(0).Item("table_type_no").ToString
                        Catch ex As Exception

                        End Try
                        Try
                            Table_Average_1.SelectedValue = ds.Tables(5).Rows(0).Item("table_average1").ToString
                        Catch ex As Exception

                        End Try
                        Try
                            Table_Average_2.SelectedValue = ds.Tables(5).Rows(0).Item("table_average2").ToString
                        Catch ex As Exception

                        End Try
                        Try
                            Table_Average_3.SelectedValue = ds.Tables(5).Rows(0).Item("table_average3").ToString
                        Catch ex As Exception

                        End Try

                        txtTableBannerHeight.Text = ds.Tables(5).Rows(0).Item("table_banner_height").ToString
                        txtUniqueID.Text = ds.Tables(5).Rows(0).Item("unique_id").ToString
                        txtGroupID.Text = ds.Tables(5).Rows(0).Item("group_id").ToString
                        txtGroupName.Text = ds.Tables(5).Rows(0).Item("group_name").ToString
                    Catch ex As Exception

                    End Try
                End If
            End If
        End If

        myCommand.Dispose()
        GenUtils.UndoConn(SQLConn)
    End Sub
    Sub LoadStubLookupsAndRecord()
        btnStubDelete.Enabled = True

        Dim ds As DataSet = New DataSet
        Dim SQLConn As SqlConnection = GenUtils.DoConn("SourceDBConnString")

        If IsNothing(Session("stub_table_id")) Then
            Session("stub_table_id") = 0
        End If

        If Session("stub_table_id") = 0 Then

            lblReportStubID.Text = "0"
            lblStubDeliverableID.Text = Session("ReportName")

        End If

        Dim sAnswerVal As String = ""

        Dim myCommand As New SqlCommand("frontend_editable_grid_stubs_lookups", SQLConn)
        myCommand.CommandType = CommandType.StoredProcedure
        myCommand.CommandTimeout = 3600

        myCommand.Parameters.Add("@domain_instance_id", SqlDbType.Int).Value = IIf(GenUtils.GetContextID("DomainIID").Length > 0, GenUtils.GetContextID("DomainIID"), 0)
        myCommand.Parameters.Add("@table_id", SqlDbType.Int).Value = Session("stub_table_id")

        Dim myData As New SqlDataAdapter(myCommand)
        myData.Fill(ds)

        If ds.Tables.Count > 0 Then

            lstFormatType2.DataSource = ds.Tables(1)
            lstFormatType2.DataTextField = "lookup_desc"
            lstFormatType2.DataValueField = "lookup_value"
            lstFormatType2.DataBind()

            dra_attr_name.DataSource = ds.Tables(0)
            dra_attr_name.DataTextField = "dvname"
            dra_attr_name.DataValueField = "derived_variable_name"
            dra_attr_name.DataBind()
            dra_attr_name.Items.Insert(0, New ListItem("Select One", ""))

            If Session("stub_table_id") = 0 Then
                btnStubDelete.Enabled = False

                txtSortAttrName.Text = ""
                txtSortAttrValue.Text = ""
                txtAttrNameOverride.Text = ""
                Stub_Super_Header.Text = ""
                txtNameValueFlag.Text = ""
                Stub_Control_Number.Text = ""

                Try
                    dra_attr_name.SelectedIndex = 0
                Catch ex As Exception

                End Try
                DRAAttrValueHolder.InnerHtml = ""
            End If

            If ds.Tables.Count > 2 Then
                'This means we're loading an existing record too
                If ds.Tables(2).Rows.Count > 0 Then
                    sAnswerVal = ds.Tables(2).Rows(0).Item("dra_attr_value").ToString
                End If
                If ds.Tables(3).Rows.Count > 0 And ds.Tables.Count > 1 Then
                    Dim sDrop As String = "<select id=""Stub_DRA_Attr_Value"" name=""Stub_DRA_Attr_Value"" style=""width:450px"">"
                    Dim s As New StringBuilder

                    For Each oRow As DataRow In ds.Tables(3).Rows
                        If oRow("answer_text").ToString = sAnswerVal Then
                            s.AppendLine("<option selected=""selected"" value=""" & sAnswerVal & """>" & sAnswerVal & "</option>")
                        Else
                            s.AppendLine("<option value=""" & oRow("answer_text").ToString & """>" & oRow("dvaname").ToString & "</option>")
                        End If
                    Next
                    s.AppendLine("</select>")
                    DRAAttrValueHolder.InnerHtml = sDrop & s.ToString

                End If

                If ds.Tables(2).Rows.Count > 0 Then
                    Try
                        lblReportStubID.Text = ds.Tables(2).Rows(0).Item("report_stub_id").ToString
                        lblStubDeliverableID.Text = ds.Tables(2).Rows(0).Item("Deliverable_ID").ToString
                        Stub_Stub_ID.Text = ds.Tables(2).Rows(0).Item("stub_id").ToString

                        txtSortAttrName.Text = ds.Tables(2).Rows(0).Item("sort_attr_name").ToString
                        txtSortAttrValue.Text = ds.Tables(2).Rows(0).Item("sort_attr_value").ToString
                        txtAttrNameOverride.Text = ds.Tables(2).Rows(0).Item("attr_name_override").ToString
                        Stub_Super_Header.Text = ds.Tables(2).Rows(0).Item("Super_Header").ToString
                        txtNameValueFlag.Text = ds.Tables(2).Rows(0).Item("name_value_flag").ToString
                        Stub_Control_Number.Text = ds.Tables(2).Rows(0).Item("Control_Column").ToString
                        txtAttr_Value_Override.Text = ds.Tables(2).Rows(0).Item("attr_value_override").ToString
                        Try
                            dra_attr_name.SelectedValue = ds.Tables(2).Rows(0).Item("dra_attr_name").ToString
                        Catch ex As Exception

                        End Try
                        Try
                            lstFormatType2.SelectedValue = ds.Tables(2).Rows(0).Item("format_type").ToString
                        Catch ex As Exception

                        End Try

                    Catch ex As Exception

                    End Try
                End If
            End If
        End If

        myCommand.Dispose()
        GenUtils.UndoConn(SQLConn)

    End Sub
    Protected Sub btnRSave_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnRSave.Click
        Dim sMKTSEG_value As String = ""
        Try
            sMKTSEG_value = Request.Form("MKTSEG_Value").ToString
        Catch ex As Exception

        End Try
        If txtReportFooter1.Text.Length > 1500 Or txtReportFooter2.Text.Length > 1500 Or txtReportFooter3.Text.Length > 1500 Then
            RadWindowManager1.RadAlert("Footer length is greater than 1500 characters for: " & _
                                       IIf(txtReportFooter1.Text.Length > 1500, "<br>Footer 1", "") & _
                                       IIf(txtReportFooter2.Text.Length > 1500, "<br>Footer 2", "") & _
                                       IIf(txtReportFooter3.Text.Length > 1500, "<br>Footer 3", "") & "<br>Save operation aborted to avoid Netezza errors.", 400, 200, "Error", Nothing)

            Exit Sub
        End If
        Param_Report_Functions.SaveReportTable(lblTableID.Text, lblDeliverableID.Text, txtTopN.Text, txtReportTableNo.Text, txtReportName.Text, txtReportFooter1.Text, _
                                               txtReportFooter2.Text, txtReportFooter3.Text, MKTSEG_Name.SelectedValue, sMKTSEG_value, txtMSDisplay.Text, _
                                               txtRankType.Text, txtSortYear.Text, txtStubID.Text, Table_Type_No.SelectedValue, Table_Average_1.SelectedValue, Table_Average_2.SelectedValue, Table_Average_3.SelectedValue, _
                                               txtTableBannerHeight.Text, Session("ProductIName"), Session("ProductIID"), txtUniqueID.Text, txtGroupID.Text, txtGroupName.Text)

        RadTabStrip1.SelectedIndex = 0
        Dim oTab As RadTab = RadTabStrip1.FindTabByText("New Table")
        oTab.Selected = False

        RadMultiPage1.SelectedIndex = 0 'Reports Grid
        RadGridReports.Rebind()

    End Sub

    Protected Sub btnRCancel_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnRCancel.Click
        RadTabStrip1.SelectedIndex = 0
        Dim oTab As RadTab = RadTabStrip1.FindTabByText("New Table")
        oTab.Selected = False
        RadMultiPage1.SelectedIndex = 0 'Reports Grid
    End Sub

    Protected Sub btnRDelete_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnRDelete.Click
        'Run delete, go back to Grid
        Dim SQLConn As SqlConnection = GenUtils.DoConn("RPConnString")

        Dim myCommand As New SqlCommand("delete_Rep_Table_Rec", SQLConn)
        myCommand.CommandType = CommandType.StoredProcedure
        myCommand.CommandTimeout = 3600

        myCommand.Parameters.Add("@table_id", SqlDbType.Int).Value = Session("report_table_id")
        myCommand.Parameters.Add("@username", SqlDbType.NVarChar, 60).Value = MProfile.UsersName

        myCommand.ExecuteNonQuery()

        myCommand.Dispose()
        GenUtils.UndoConn(SQLConn)

        RadTabStrip1.SelectedIndex = 0
        Dim oTab As RadTab = RadTabStrip1.FindTabByText("New Table")
        oTab.Selected = False
        RadMultiPage1.SelectedIndex = 0 'Reports Grid
        RadGridReports.Rebind()

    End Sub
    Sub EditReportItem(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs)
        'Dim mygridItem As GridItem = CType(CType(sender, ImageButton).Parent.Parent, GridItem)
        'Dim TableID As String = mygridItem.Cells(3).Text
        Dim dataItem As GridDataItem = CType(sender.Parent.Parent, Telerik.Web.UI.GridDataItem)
        Dim TableID As String = dataItem.OwnerTableView.DataKeyValues(dataItem.ItemIndex)("Table_ID")

        Session("report_table_id") = TableID
        Dim oTab As RadTab = RadTabStrip1.FindTabByText("New Table")
        oTab.Selected = True
        RadMultiPage1.SelectedIndex = 4
        LoadReportLookupsAndRecord()
    End Sub
    Sub EditBannerItem(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs)
        'Dim mygridItem As GridItem = CType(CType(sender, ImageButton).Parent.Parent, GridItem)
        'Dim TableID As String = mygridItem.Cells(3).Text
        Dim dataItem As GridDataItem = CType(sender.Parent.Parent, Telerik.Web.UI.GridDataItem)
        Dim TableID As String = dataItem.OwnerTableView.DataKeyValues(dataItem.ItemIndex)("Table_ID")

        Session("banner_table_id") = TableID
        Dim oTab As RadTab = RadTabStrip1.FindTabByText("New Banner")
        oTab.Selected = True
        RadMultiPage1.SelectedIndex = 5
        LoadBannerLookupsAndRecord()
    End Sub
    Sub EditStubItem(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs)
        'Dim mygridItem As GridItem = CType(CType(sender, ImageButton).Parent.Parent, GridItem)
        'Dim TableID As String = mygridItem.Cells(3).Text
        Dim dataItem As GridDataItem = CType(sender.Parent.Parent, Telerik.Web.UI.GridDataItem)
        Dim TableID As String = dataItem.OwnerTableView.DataKeyValues(dataItem.ItemIndex)("Table_ID")

        Session("stub_table_id") = TableID
        Dim oTab As RadTab = RadTabStrip1.FindTabByText("New Stub")
        oTab.Selected = True
        RadMultiPage1.SelectedIndex = 6
        LoadStubLookupsAndRecord()
    End Sub

    Protected Sub btnBannerSave_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnBannerSave.Click
        Dim sdra_value As String = ""
        Try
            sdra_value = Request.Form("dra_value").ToString
        Catch ex As Exception

        End Try
        Param_Report_Functions.SaveBannerTable(Session("banner_table_id"), Session("ReportID"), banner_report_table_no.SelectedValue, txtBannerDisplay.Text, txtOrderBanner.Text, _
                                               txtNumOfYears.Text, txtSuperHeader.Text, Source_System_Trend_ID.SelectedValue, DRA_Name.SelectedValue, sdra_value, txtIndicator.Text, txtObjectIndicator.Text, Numerator.SelectedValue, _
                                               Denominator.SelectedValue, Number_Format.SelectedValue, Decimals.SelectedValue, txtControlColumn.Text, Session("ProductIName"), Session("ProductIID"), lstFormatType.SelectedValue)

        RadTabStrip1.SelectedIndex = 1
        Dim oTab As RadTab = RadTabStrip1.FindTabByText("New Banner")
        oTab.Selected = False

        RadMultiPage1.SelectedIndex = 1 'Banner Grid
        RadGridBanners.Rebind()
    End Sub

    Protected Sub btnBannerDelete_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnBannerDelete.Click
        Dim SQLConn As SqlConnection = GenUtils.DoConn("PRConnString")

        Dim myCommand As New SqlCommand("delete_Banner_Table_Rec", SQLConn)
        myCommand.CommandType = CommandType.StoredProcedure
        myCommand.CommandTimeout = 3600

        myCommand.Parameters.Add("@banner_id", SqlDbType.Int).Value = Session("banner_table_id")
        myCommand.Parameters.Add("@username", SqlDbType.NVarChar, 60).Value = MProfile.UsersName

        myCommand.ExecuteNonQuery()

        myCommand.Dispose()
        GenUtils.UndoConn(SQLConn)

        RadTabStrip1.SelectedIndex = 1
        Dim oTab As RadTab = RadTabStrip1.FindTabByText("New Banner")
        oTab.Selected = False

        RadMultiPage1.SelectedIndex = 1 'Banner Grid
        RadGridBanners.Rebind()
    End Sub

    Protected Sub btnBannerCancel_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnBannerCancel.Click
        RadTabStrip1.SelectedIndex = 0
        Dim oTab As RadTab = RadTabStrip1.FindTabByText("New Banner")
        oTab.Selected = False
        RadTabStrip1.SelectedIndex = 1
        RadMultiPage1.SelectedIndex = 1 'Banner Grid
    End Sub

    Protected Sub btnStubSave_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnStubSave.Click
        Dim sStub_DRA_Attr_Value As String = ""
        Try
            sStub_DRA_Attr_Value = Request.Form("stub_dra_attr_value").ToString
        Catch ex As Exception

        End Try
        Param_Report_Functions.SaveStubTable(Session("stub_table_id"), Session("ReportID"), Stub_Stub_ID.Text, dra_attr_name.SelectedValue, sStub_DRA_Attr_Value, txtSortAttrName.Text, _
                                             txtSortAttrValue.Text, txtAttrNameOverride.Text, Stub_Super_Header.Text, txtNameValueFlag.Text, Stub_Control_Number.Text, txtAttr_Value_Override.Text, lstFormatType2.SelectedValue)

        RadTabStrip1.SelectedIndex = 2
        Dim oTab As RadTab = RadTabStrip1.FindTabByText("New Stub")
        oTab.Selected = False

        RadMultiPage1.SelectedIndex = 2 'Stub Grid
        RadGridStubs.Rebind()
    End Sub

    Protected Sub btnStubCancel_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnStubCancel.Click
        RadTabStrip1.SelectedIndex = 2
        Dim oTab As RadTab = RadTabStrip1.FindTabByText("New Stub")
        oTab.Selected = False

        RadMultiPage1.SelectedIndex = 2 'Stub Grid
        RadGridStubs.Rebind()
    End Sub

    Protected Sub btnStubDelete_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnStubDelete.Click
        Dim SQLConn As SqlConnection = GenUtils.DoConn("PRConnString")

        Dim myCommand As New SqlCommand("delete_Stub_Table_Rec", SQLConn)
        myCommand.CommandType = CommandType.StoredProcedure
        myCommand.CommandTimeout = 3600

        myCommand.Parameters.Add("@report_stub_id", SqlDbType.Int).Value = Session("stub_table_id")
        myCommand.Parameters.Add("@username", SqlDbType.NVarChar, 60).Value = MProfile.UsersName

        myCommand.ExecuteNonQuery()

        myCommand.Dispose()
        GenUtils.UndoConn(SQLConn)

        RadTabStrip1.SelectedIndex = 2
        Dim oTab As RadTab = RadTabStrip1.FindTabByText("New Stub")
        oTab.Selected = False

        RadMultiPage1.SelectedIndex = 2 'Stub Grid
        RadGridStubs.Rebind()
    End Sub
    Sub LoadTree()
        radTreeResort.Nodes.Clear()

        Dim SQLConn As SqlConnection = GenUtils.DoConn("PRConnString")
        Dim DS As DataSet = New DataSet

        Dim myCommand As New SqlCommand("select table_id, '<b>' + cast(report_table_no as nvarchar(25)) + '</b> - ' + Report_Name  as table_text from a_report_table where deliverable_id = " & Session("ReportID") & _
                                        " order by report_table_no", SQLConn)
        myCommand.CommandType = CommandType.Text
        myCommand.CommandTimeout = 3600

        Dim myReader As New SqlDataAdapter(myCommand)
        myReader.Fill(DS)

        radTreeResort.DataSource = DS.Tables(0)
        radTreeResort.DataFieldID = "table_id"
        radTreeResort.DataValueField = "table_id"

        radTreeResort.DataTextField = "table_text"
        radTreeResort.DataBind()

        Dim sXML As String = ""

        For Each oNode As RadTreeNode In radTreeResort.GetAllNodes
            sXML &= oNode.Index + 1 & "|" & oNode.Value & "~"
        Next oNode

        CurrentOrder.Value = sXML

        myCommand.Dispose()
        GenUtils.UndoConn(SQLConn)

    End Sub

    Protected Sub btnResort_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnSortReports.Click
        Dim sOut As String = CurrentOrder.Value
        sOut.Remove(sOut.Length - 1, 1) 'remove trailing "~"

        Dim sTemp() As String = Split(sOut, "~")
        Dim sXML As String = "<columndata>"

        For i As Int16 = 0 To UBound(sTemp) - 1
            Dim sTemp2() As String = Split(sTemp(i), "|")

            sXML &= "<row id=""" & sTemp2(0) & """ value=""" & sTemp2(1) & """ />" & vbCrLf

        Next i

        sXML &= "</columndata>"

        Dim SQLConn As SqlConnection = GenUtils.DoConn("SourceDBConnString")

        Dim myCommand As New SqlCommand("editable_grid_ReSort_Report_Table", SQLConn)
        myCommand.CommandType = CommandType.StoredProcedure
        myCommand.CommandTimeout = 3600
        myCommand.Parameters.Add("@input", SqlDbType.Xml).Value = sXML

        myCommand.ExecuteNonQuery()

        myCommand.Dispose()
        GenUtils.UndoConn(SQLConn)

        LoadTree()
        RadGridReports.Rebind()
    End Sub

    Protected Sub btnResortCancel_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnResortCancel.Click
        RadTabStrip1.SelectedIndex = 0
        Dim oTab As RadTab = RadTabStrip1.FindTabByText("Re-sort Tables")
        oTab.Selected = False
        RadMultiPage1.SelectedIndex = 0 'Reports Grid
        RadGridReports.Rebind()

    End Sub

End Class