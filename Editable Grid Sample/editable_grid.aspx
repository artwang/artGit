<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/MainInterface_NoUpPanel.Master" CodeBehind="editable_grid.aspx.vb" Inherits="myProject.editable_grid" %>
<%@ Register assembly="Telerik.Web.UI" namespace="Telerik.Web.UI" tagprefix="telerik" %>

<asp:Content ID="Content1" ContentPlaceHolderID="ContentPlaceHolder1" Runat="Server">

    <script language="javascript" type="text/javascript">
        function MKSSelected(sDrop, ReturnDiv, fieldname) {
            var i = sDrop.selectedIndex;
        
            var sText = sDrop.options[i].text;
            var sVal = sDrop.options[i].value;
        
            var sDV = sText.replace(" - " + sVal, "");
            sDV = sDV.replace("DVID: ", "");
        
            myProject.myService.GetDVAnswers(sDV, ReturnDiv, fieldname, OnDVLookupComplete, OnDVError);
        }

        function OnDVLookupComplete(result) {
            var oRes = result.split("~");
        
            var oLit = document.getElementById(oRes[1]);
            oLit.innerHTML = oRes[0];
        }
        
        
        function OnDVError(result) {
            alert("Error: " + result.get_message());
        }
        
        

        function GoSave(oInput) {
            oInput.style.border = "1px solid blue";
            var Validation = oInput.id.split("|");

            if (oInput.id.indexOf("Footer") > 0) {
                if (oInput.value.length > 1500) {
                    oInput.value = oInput.value.substring(0, 1500);
                    alert("Footer text is longer than 1500 characters.  Data has been truncated to avoid Netezza errors.");
                }
            }
        
            if (Validation[Validation.length - 2] == 'N') {
                if (isNaN(oInput.value)) { //is not a number
                    alert("This field accepts numbers only.  Operation aborted.");
                    oInput.value = Validation[Validation.length - 1];
                    return false;
                }
            }
        
            //call web service here, pass the data, and wait for response!
            myProject.myService.SaveInput(oInput.id, oInput.value, OnUpdateComplete, OnUpdateError);
        }
        function ChangeBorder(oInput) {
            oInput.style.border = "1px solid red";
        }
        
        function OnUpdateComplete(result) {
            var sres = result.split("~");
            var sInput = sres[0];
            var sstatus = sres[1];
            var sOut = sres[2];
            var oInput = document.getElementById(sInput)
            var sInp = sInput.split("|");
        
            //table_id|123|Report_Table_No|a_Report_table|NorS|OldNumValue  (Nors = Numeric or string)
            //    0     1         2               3        4      5
            if (sstatus == 'OK') {
                if (sInp[4] == 'D') { //special handling for drop-down
                    sInp[4] = 'S';
                    var oDiv = document.getElementById(sInput.replace('|D|', '|S|'));
                    oDiv.style.border = '1px solid blue';
                    oDiv.id = sInp[0] + '|' + sInp[1] + '|' + sInp[2] + '|' + sInp[3] + '|S|' + oInput.value;
                    oInput.id = sInp[0] + '|' + sInp[1] + '|' + sInp[2] + '|' + sInp[3] + '|D|' + oInput.value;
                    oDiv.setAttribute('onclick', 'GetDropvalues(this);')
                }
                else {
                    oInput.style.border = '1px solid blue';
                    oInput.id = sInp[0] + '|' + sInp[1] + '|' + sInp[2] + '|' + sInp[3] + '|' + sInp[4] + '|' + oInput.value;
                }
            }
            else {
                oInput.style.border = '1px solid red';
                oInput.value = sOut;
            }        
        }
        
        
        function OnUpdateError(result) {
            alert("Error: " + result.get_message());
        }
        
        function GetDropValues(oDrop) {
            //table_id|123|Report_Table_No|a_Report_table|NorS|OldNumValue  (Nors = Numeric or string)
            //    0     1         2               3        4      5
            myProject.myService.GetDropDownValues(oDrop.id, OnDropLookupComplete, OnDropError);
        }

        function GetTrendDropValues(oDrop) {
            myProject.myService.GetTrendDropDownValues(oDrop.id, OnDropLookupComplete, OnDropError);
        }

        function OnDropLookupComplete(result) {
            var oRes = result.split("~");
            var oDiv = document.getElementById(oRes[1]);
            oDiv.innerHTML = oRes[0];
            oDiv.style.border = "1px solid red";
            oDiv.onclick = null;
        }
        
        
        function OnDropError(result) {
            alert("Error: " + result.get_message());
        }
        
        function GetDVDropValues(oDrop) {
            var sDrop = document.getElementById(oDrop.id.replace("|S|", "|D|"));

            var i = sDrop.selectedIndex;
        
            var sText = sDrop.options[i].text;
            var sVal = sDrop.options[i].value;
        
            var sDV = sText.replace(" - " + sVal, "");
            sDV = sDV.replace("DVID: ", "");

            myProject.myService.GetDVQuestions(oDrop.id, sText, OnDropLookupComplete, OnDVError);
        }
        
        function GetDVADropValues(oDrop) {
            var sDrop = document.getElementById(oDrop.id.replace("|S|", "|D|"));
        
            var i = sDrop.selectedIndex;

            var sVal = sDrop.options[i].value;

            myProject.myService.GetDVAQuestions(oDrop.id, sVal, OnDropLookupComplete, OnDVError);
        }
        
        function onDVMultiLookupComplete(result) {
            var oRes = result.split("~");
            var oDiv = document.getElementById(oRes[1]);
            oDiv.innerHTML = oRes[0];
            oDiv.style.border = "1px solid red";
            oDiv.onclick = null;
        }
        
        function GoSaveDV(oInput) {
            // Save DV value.. then load DVA values into correct drop-down.
            //call web service here, pass the data, and wait for response!
            myProject.myService.SaveInputDV(oInput.id, oInput.value, OnUpdateDVComplete, OnUpdateError);
        }
        
        function OnUpdateDVComplete(result) {
            /*
            InputID & "~" & sRtn & "~" &  NewID & "~" & s.ToString & table_row_count
            '         0              1            2                 3              4
            '------------------------------------------------------------------------
            */
            var oRes = result.split("~");
            if (oRes[1] == "OK") {
                var oDiv = document.getElementById(oRes[0].replace("|D|", "|S|"));
                oDiv.style.border = "1px solid blue";
                oDiv.setAttribute('onclick', 'GetDVDropValues(this);');
        
                var oAnsDrop = document.getElementById(oRes[2].replace("|D|", "|S|"));
                oAnsDrop.innerHTML = oRes[3];
                if (oRes[4] > 1) {
                    oAnsDrop.style.border = "1px solid red";
                    oAnsDrop.onclick = null;
                }
                else {
                    oAnsDrop.style.border = "1px solid blue";
                    oAnsDrop.onclick = null;
                }
            }
        }
        
        function ClientNodeDropping(sender, eventArgs) {
            if (eventArgs.get_sourceNode().get_level() != eventArgs.get_destNode().get_level()) {
                alert("You cannot drag nodes between levels");
                eventArgs.set_cancel(true);
            }
        
            var tree = $find("<%= radTreeResort.ClientID %>");
            //tree.trackChanges();
            // var node = eventArgs.get_sourceNode();  //new Telerik.Web.UI.RadTreeNode();
        
            var destIndex = eventArgs.get_destNode().get_index();
            var node = new Telerik.Web.UI.RadTreeNode();
            node.set_text('<b>' + eventArgs.get_sourceNode().get_text() + '</b>');
            node.set_value(eventArgs.get_sourceNode().get_value());
        
            tree.get_nodes().insert(destIndex, node);
        
            tree.get_nodes().remove(eventArgs.get_sourceNode());
            tree.commitChanges();
        
            var sOrder = "";
        
            for (var i = 0; i < tree.get_nodes().get_count(); i++) {
                var onode = tree.get_nodes().getNode(i);
                sOrder += (i + 1) + "|" + onode.get_value() + "~";
            }

            document.getElementById("<%= CurrentOrder.ClientID %>").value = sOrder;        
        }
        
    </script>

</asp:Content>

<asp:Content ID="Content2" ContentPlaceHolderID="ContentPlaceHolder3" Runat="Server">
    <telerik:RadAjaxLoadingPanel ID="RadAjaxLoadingPanel1" runat="server" Height="75px" Width="75px">
        <img alt="Loading..." src='<%= RadAjaxLoadingPanel.GetWebResourceUrl(Page, "Telerik.Web.UI.Skins.Default.Ajax.loading.gif") %>' style="border: 0px;" />
    </telerik:RadAjaxLoadingPanel>
    <telerik:RadAjaxPanel ID="RadAjaxPanel1" runat="server" LoadingPanelID="RadAjaxLoadingPanel1" Height="100%">
        
        <asp:HiddenField runat="server" ID="CurrentOrder" />
        <telerik:RadWindowManager ID="RadWindowManager1" runat="server" Skin="Office2007" Behaviors="Default" EnableViewState="false"></telerik:RadWindowManager>
        <telerik:RadTabStrip ID="RadTabStrip1" runat="server" Skin="Office2007" MultiPageID="RadMultiPage1" SelectedIndex="0" Height="100%">
            <Tabs>
                <telerik:RadTab runat="server" Text="Tables" Value="Reports" PageViewID="ReportsPage">
                    <Tabs>
                        <telerik:RadTab runat="server" Text="New Table" Value="Reports_New_Report" PageViewID="NewReport"></telerik:RadTab>
                        <telerik:RadTab runat="server" Text="Re-sort Tables" Value="Reports_Resort" PageViewID="ResortReports"></telerik:RadTab>
                    </Tabs>
                </telerik:RadTab>
                <telerik:RadTab runat="server" Text="Banners" Value="Banners" PageViewID="BannersPage">
                    <Tabs>
                        <telerik:RadTab runat="server" Text="New Banner" Value="Banners_New_Report" PageViewID="NewBanner"></telerik:RadTab>
                    </Tabs>
                </telerik:RadTab>
                <telerik:RadTab runat="server" Text="Stubs" Value="Stubs" PageViewID="StubsPage">
                    <Tabs>
                        <telerik:RadTab runat="server" Text="New Stub" Value="Stubs_New_Report" PageViewID="NewStub"></telerik:RadTab>
                    </Tabs>
                </telerik:RadTab>
                <telerik:RadTab runat="server" Text="Import / Export" ImageUrl="~/images/excel.gif" Value="Export" PageViewID="pageImportExport">
                    <Tabs>
                        <telerik:RadTab runat="server" Text="Import" Value="Import"></telerik:RadTab>
                    </Tabs>
                </telerik:RadTab>
            </Tabs>
        </telerik:RadTabStrip>
        <telerik:RadMultiPage ID="RadMultiPage1" Runat="server" SelectedIndex="0" Height="100%">
            <telerik:RadPageView ID="ReportsPage" runat="server">
                <telerik:RadGrid ID="RadGridReports" runat="server" AllowFilteringByColumn="True" AllowPaging="True" AllowSorting="True" GridLines="Horizontal" Skin="Vista">
                    <GroupingSettings CaseSensitive="false" />
                    <ClientSettings>
                        <Scrolling AllowScroll="false" UseStaticHeaders="True" />
                    </ClientSettings>
                    <MasterTableView AutoGenerateColumns="False" CellSpacing="0" DataKeyNames="Table_ID" EditMode="PopUp" PageSize="10" Width="80%">
                        <Columns>
                            <telerik:GridTemplateColumn AllowFiltering="false" HeaderText="Edit" UniqueName="TemplateColumn">
                                <ItemTemplate>
                                    <div align="left">
                                        <asp:ImageButton ID="ReportEditButton" runat="server" ImageUrl="~/images/edit.gif" OnClick="EditReportItem" ToolTip="Edit Item" style="cursor:pointer;" />
                                    </div>
                                </ItemTemplate>
                                <ItemStyle Wrap="false" />
                            </telerik:GridTemplateColumn>
                            <telerik:GridBoundColumn DataField="Table_ID" DataType="System.Int32" HeaderText="Table_ID" ReadOnly="True" SortExpression="Table_ID" UniqueName="Table_ID" Visible="False">
                            </telerik:GridBoundColumn>
                            <telerik:GridBoundColumn DataField="Deliverable_ID" DataType="System.Int32" HeaderText="Deliverable_ID" SortExpression="Deliverable_ID" UniqueName="Deliverable_ID" Visible="False">
                            </telerik:GridBoundColumn>
                            <telerik:GridBoundColumn DataField="Product_ID" DataType="System.Int32" HeaderText="Product_ID" SortExpression="Product_ID" UniqueName="Product_ID" Visible="False">
                            </telerik:GridBoundColumn>
                            <telerik:GridBoundColumn DataField="Product_Instance_Name" HeaderText="Product_Instance_Name" SortExpression="Product_Instance_Name" UniqueName="Product_Instance_Name" Visible="False">
                            </telerik:GridBoundColumn>
                            <telerik:GridTemplateColumn DataField="Report_Table_No" DataType="System.Decimal" HeaderText="Report_Table_No" SortExpression="Report_Table_No" UniqueName="Report_Table_No">
                                <ItemStyle Wrap="False" />
                                <itemtemplate>
                                <%#Eval("Report_Table_No")%>
                                </itemtemplate>
                            </telerik:GridTemplateColumn>
                            <telerik:GridTemplateColumn DataField="TopN" DataType="System.Decimal" HeaderText="TopN" SortExpression="TopN" UniqueName="TopN">
                                <ItemTemplate>
                                    <input ID='table_id|<%# Eval("table_id") & "|TopN|a_Report_table|N|" & Eval("TopN") %>' onblur="GoSave(this);" onfocus="ChangeBorder(this);" onclick="ChangeBorder(this);" style="width: 100px;font-family: tahoma; font-size: 8pt; border: 1px solid blue; text-align: right;" type="text" value='<%#Eval("TopN")%>' />
                                </ItemTemplate>
                                <ItemStyle Wrap="False" />
                            </telerik:GridTemplateColumn>
                            <telerik:GridTemplateColumn DataField="Report_Name" HeaderText="Report_Name" SortExpression="Report_Name" UniqueName="Report_Name">
                                <ItemTemplate>
                                    <textarea ID='table_id|<%# Eval("table_id") & "|Report_Name|a_Report_table|S|" & Server.URLEncode(Eval("Report_Name").ToString.replace("'", "\'").replace("""", "\""")) %>' onblur="GoSave(this);" onfocus="ChangeBorder(this);" onclick="ChangeBorder(this);" rows="2" style="width: 450px; font-family: tahoma; font-size: 8pt; height: 25px; border: 1px solid blue; text-align: left;"><%#Eval("Report_Name")%></textarea>
                                </ItemTemplate>
                                <ItemStyle Wrap="False" />
                            </telerik:GridTemplateColumn>
                            <telerik:GridTemplateColumn DataField="Report_Footer" HeaderText="Report_Footer" SortExpression="Report_Footer" UniqueName="Report_Footer">
                                <ItemTemplate>
                                    <textarea ID='table_id|<%# Eval("table_id") & "|Report_Footer|a_Report_table|S|0" %>' onblur="GoSave(this);" onfocus="ChangeBorder(this);" onclick="ChangeBorder(this);" rows="2" style="width: 450px; font-family: tahoma; font-size: 8pt; height: 25px; border: 1px solid blue; text-align: left;"><%#Eval("Report_Footer")%></textarea>
                                </ItemTemplate>
                                <ItemStyle Wrap="False" />
                            </telerik:GridTemplateColumn>
                            <telerik:GridTemplateColumn DataField="Report_Footer2" HeaderText="Report_Footer2" SortExpression="Report_Footer2" UniqueName="Report_Footer2">
                                <ItemTemplate>
                                    <textarea ID='table_id|<%# Eval("table_id") & "|Report_Footer2|a_Report_table|S|0" %>' onblur="GoSave(this);" onfocus="ChangeBorder(this);" onclick="ChangeBorder(this);" rows="2" style="width: 450px; font-family: tahoma; font-size: 8pt; height: 25px; border: 1px solid blue; text-align: left;"><%#Eval("Report_Footer2")%></textarea>
                                </ItemTemplate>
                                <ItemStyle Wrap="False" />
                            </telerik:GridTemplateColumn>
                            <telerik:GridTemplateColumn DataField="Report_Footer3" HeaderText="Report_Footer3" SortExpression="Report_Footer3" UniqueName="Report_Footer3">
                                <ItemTemplate>
                                    <textarea ID='table_id|<%# Eval("table_id") & "|Report_Footer3|a_Report_table|S|0" %>' onblur="GoSave(this);" onfocus="ChangeBorder(this);" onclick="ChangeBorder(this);" rows="2" style="width: 450px; font-family: tahoma; font-size: 8pt; height: 25px; border: 1px solid blue; text-align: left;"><%#Eval("Report_Footer3")%></textarea>
                                </ItemTemplate>
                                <ItemStyle Wrap="False" />
                            </telerik:GridTemplateColumn>
                            <telerik:GridTemplateColumn DataField="MKTSEG_Name" HeaderText="MKTSEG_Name" SortExpression="MKTSEG_Name" UniqueName="MKTSEG_Name">
                                <ItemStyle Wrap="False" />
                                <ItemTemplate>
                                    <div ID='table_id|<%# Eval("table_id") & "|MKTSEG_Name|a_Report_table|S|0" %>' onclick="GetDVDropValues(this);" style="width: 450px; border: 1px solid blue;">
                                        <select ID='table_id|<%# Eval("table_id") & "|MKTSEG_Name|a_Report_table|D|0" %>' onchange="GoSaveDV(this);" style="width: 450px; font-family: tahoma; font-size: 8pt;">
                                            <option selected="selected" value='<%# Eval("MKTSEG_Name") %>'><%#Eval("MKTSEG_Name")%></option>
                                        </select>
                                    </div>
                                </ItemTemplate>
                            </telerik:GridTemplateColumn>
                            <telerik:GridTemplateColumn DataField="MKTSEG_Value" HeaderText="MKTSEG_Value" SortExpression="MKTSEG_Value" UniqueName="MKTSEG_Value">
                                <ItemTemplate>
                                    <div ID='table_id|<%# Eval("table_id") & "|MKTSEG_Value|a_Report_table|S|0" %>' onclick="GetDVADropValues(this);" style="width: 450px; border: 1px solid blue;">
                                        <select ID='table_id|<%# Eval("table_id") & "|MKTSEG_Value|a_Report_table|D|0" %>' onchange="GoSave(this);" style="width: 450px; font-family: tahoma; font-size: 8pt;">
                                            <option selected="selected" value='<%# Eval("MKTSEG_Value") %>'><%#Eval("MKTSEG_Value")%></option>
                                        </select>
                                    </div>
                                </ItemTemplate>
                                <ItemStyle Wrap="False" />
                            </telerik:GridTemplateColumn>
                            <telerik:GridTemplateColumn DataField="MKTSEG_Display" HeaderText="MKTSEG_Display" SortExpression="MKTSEG_Display" UniqueName="MKTSEG_Display">
                                <ItemTemplate>
                                    <textarea ID='table_id|<%# Eval("table_id") & "|MKTSEG_Display|a_Report_table|S|0" %>' onblur="GoSave(this);" onfocus="ChangeBorder(this);" onclick="ChangeBorder(this);" rows="2" style="width: 450px; font-family: tahoma; font-size: 8pt; height: 25px; border: 1px solid blue; text-align: left;"><%#Eval("MKTSEG_Display")%></textarea>
                                </ItemTemplate>
                                <ItemStyle Wrap="False" />
                            </telerik:GridTemplateColumn>
                            <telerik:GridTemplateColumn DataField="Rank_Type" HeaderText="Rank_Type" SortExpression="Rank_Type" UniqueName="Rank_Type">
                                <ItemTemplate>
                                    <input ID='table_id|<%# Eval("table_id") & "|Rank_Type|a_Report_table|S|0" %>' maxlength="50" onblur="GoSave(this);" onfocus="ChangeBorder(this);" onclick="ChangeBorder(this);" style="width: 100px; font-family: tahoma; font-size: 8pt; border: 1px solid blue; text-align: right;" type="text" value='<%#Eval("Rank_Type")%>' />
                                </ItemTemplate>
                                <ItemStyle Wrap="False" />
                            </telerik:GridTemplateColumn>
                            <telerik:GridTemplateColumn DataField="Sort_Year" DataType="System.Decimal" HeaderText="Sort_Year" SortExpression="Sort_Year" UniqueName="Sort_Year">
                                <ItemTemplate>
                                    <input ID='table_id|<%# Eval("table_id") & "|Sort_Year|a_Report_table|N|" & Eval("Sort_Year") %>' onblur="GoSave(this);" onfocus="ChangeBorder(this);" onclick="ChangeBorder(this);" style="width: 100px; font-family: tahoma; font-size: 8pt; border: 1px solid blue; text-align: right;" type="text" value='<%#Eval("Sort_Year")%>' />
                                </ItemTemplate>
                                <ItemStyle Wrap="False" />
                            </telerik:GridTemplateColumn>
                            <telerik:GridTemplateColumn DataField="Stub_Id" DataType="System.Int32" HeaderText="Stub_Id" SortExpression="Stub_Id" UniqueName="Stub_Id">
                                <ItemTemplate>
                                    <input ID='table_id|<%# Eval("table_id") & "|Stub_Id|a_Report_table|N|" & Eval("Stub_Id") %>' onblur="GoSave(this);" onfocus="ChangeBorder(this);" onclick="ChangeBorder(this);" style="width: 100px; font-family: tahoma; font-size: 8pt; border: 1px solid blue; text-align: right;" type="text" value='<%#Eval("Stub_Id")%>' />
                                </ItemTemplate>
                                <ItemStyle Wrap="False" />
                            </telerik:GridTemplateColumn>
                            <telerik:GridTemplateColumn DataField="Table_Type_No" DataType="System.Decimal" HeaderText="Table_Type_No" SortExpression="Table_Type_No" UniqueName="Table_Type_No">
                                <ItemStyle Wrap="False" />
                                <ItemTemplate>
                                    <div ID='table_id|<%# Eval("table_id") & "|Table_Type_No|a_Report_table|S|" & Eval("Table_Type_No") %>' onclick="GetDropValues(this);" style="width: 100%; border: 1px solid blue;">
                                        <select ID='table_id|<%# Eval("table_id") & "|Table_Type_No|a_Report_table|D|" & Eval("Table_Type_No") %>' onchange="GoSave(this);" style="width: 200px; font-family: tahoma; font-size: 8pt;">
                                            <option selected="selected" value='<%# Eval("table_Type_No") %>'><%#Eval("Table_Type_NO")%></option>
                                        </select>
                                    </div>
                                </ItemTemplate>
                            </telerik:GridTemplateColumn>
                            <telerik:GridTemplateColumn DataField="Table_Average1" HeaderText="Table_Average1" SortExpression="Table_Average1" UniqueName="Table_Average1">
                                <ItemTemplate>
                                    <div ID='table_id|<%# Eval("table_id") & "|Table_Average1|a_Report_table|S|" & Eval("Table_Average1") %>' onclick="GetDropValues(this);" style="width: 100%; border: 1px solid blue;">
                                        <select ID='table_id|<%# Eval("table_id") & "|Table_Average1|a_Report_table|D|" & Eval("Table_Average1") %>' onchange="GoSave(this);" style="width: 200px; font-family: tahoma; font-size: 8pt;">
                                            <option selected="selected" value='<%# Eval("Table_Average1") %>'><%#Eval("Table_Average1")%></option>
                                        </select>
                                    </div>
                                </ItemTemplate>
                                <ItemStyle Wrap="False" />
                            </telerik:GridTemplateColumn>
                            <telerik:GridTemplateColumn DataField="Table_Average2" HeaderText="Table_Average2" SortExpression="Table_Average2" UniqueName="Table_Average2">
                                <ItemTemplate>
                                    <div ID='table_id|<%# Eval("table_id") & "|Table_Average2|a_Report_table|S|" & Eval("Table_Average2") %>' onclick="GetDropValues(this);" style="width: 100%; border: 1px solid blue;">
                                        <select ID='table_id|<%# Eval("table_id") & "|Table_Average2|a_Report_table|D|" & Eval("Table_Average2") %>' onchange="GoSave(this);" style="width: 200px; font-family: tahoma; font-size: 8pt;">
                                            <option selected="selected" value='<%# Eval("Table_Average2") %>'><%#Eval("Table_Average2")%></option>
                                        </select>
                                    </div>
                                </ItemTemplate>
                                <ItemStyle Wrap="False" />
                            </telerik:GridTemplateColumn>
                            <telerik:GridTemplateColumn DataField="Table_Average3" HeaderText="Table_Average3" SortExpression="Table_Average3" UniqueName="Table_Average3">
                                <ItemTemplate>
                                    <div ID='table_id|<%# Eval("table_id") & "|Table_Average3|a_Report_table|S|" & Eval("Table_Average3") %>' onclick="GetDropValues(this);" style="width: 100%; border: 1px solid blue;">
                                        <select ID='table_id|<%# Eval("table_id") & "|Table_Average3|a_Report_table|D|" & Eval("Table_Average3") %>' onchange="GoSave(this);" style="width: 200px; font-family: tahoma; font-size: 8pt;">
                                            <option selected="selected" value='<%# Eval("Table_Average3") %>'><%#Eval("Table_Average3")%></option>
                                        </select>
                                    </div>
                                </ItemTemplate>
                                <ItemStyle Wrap="False" />
                            </telerik:GridTemplateColumn>
                            <telerik:GridTemplateColumn DataField="Table_Banner_Height" DataType="System.Decimal" HeaderText="Table_Banner_Height" SortExpression="Table_Banner_Height" UniqueName="Table_Banner_Height">
                                <ItemTemplate>
                                    <input ID='table_id|<%# Eval("table_id") & "|Table_Banner_Height|a_Report_table|N|" & Eval("Table_Banner_Height") %>' onblur="GoSave(this);" onfocus="ChangeBorder(this);" onclick="ChangeBorder(this);" style="width: 100px; font-family: tahoma; font-size: 8pt; border: 1px solid blue; text-align: right;" type="text" value='<%#Eval("Table_Banner_Height")%>' />
                                </ItemTemplate>
                                <ItemStyle Wrap="False" />
                            </telerik:GridTemplateColumn>
                            <telerik:GridTemplateColumn DataField="Unique_ID" HeaderText="Unique_ID" SortExpression="Unique_ID" UniqueName="Unique_ID">
                                <ItemTemplate>
                                    <input ID='table_id|<%# Eval("table_id") & "|Unique_ID|a_Report_table|S|" & Eval("Unique_ID") %>' onblur="GoSave(this);" onfocus="ChangeBorder(this);" onclick="ChangeBorder(this);" style="width: 100px; font-family: tahoma; font-size: 8pt; border: 1px solid blue; text-align: right;" type="text" value='<%#Eval("Unique_ID")%>' />
                                </ItemTemplate>
                                <ItemStyle Wrap="False" />
                            </telerik:GridTemplateColumn>
                            <telerik:GridTemplateColumn DataField="Group_ID" HeaderText="Group_ID" SortExpression="Group_ID" UniqueName="Group_ID">
                                <ItemTemplate>
                                    <input ID='table_id|<%# Eval("table_id") & "|Group_ID|a_Report_table|S|" & Eval("Group_ID") %>' onblur="GoSave(this);" onfocus="ChangeBorder(this);" onclick="ChangeBorder(this);" style="width: 100px; font-family: tahoma; font-size: 8pt; border: 1px solid blue; text-align: right;" type="text" value='<%#Eval("Group_ID")%>' />
                                </ItemTemplate>
                                <ItemStyle Wrap="False" />
                            </telerik:GridTemplateColumn>
                            <telerik:GridTemplateColumn DataField="Group_Name" HeaderText="Group_Name" SortExpression="Group_Name" UniqueName="Group_Name">
                                <ItemTemplate>
                                    <textarea ID='table_id|<%# Eval("table_id") & "|Group_Name|a_Report_table|S|0" %>' onblur="GoSave(this);" onfocus="ChangeBorder(this);" onclick="ChangeBorder(this);" rows="2" style="width: 450px; font-family: tahoma; font-size: 8pt; height: 25px; border: 1px solid blue; text-align: left;"><%#Eval("Group_Name")%></textarea>
                                </ItemTemplate>
                                <ItemStyle Wrap="False" />
                            </telerik:GridTemplateColumn>
                        </Columns>
                        <PagerStyle Mode="NextPrevNumericAndAdvanced" Position="TopAndBottom" />
                    </MasterTableView>
                    <AlternatingItemStyle BackColor="#E0E0E0" />
                </telerik:RadGrid>
            </telerik:RadPageView>
            <telerik:RadPageView ID="BannersPage" runat="server">
                <telerik:RadGrid ID="RadGridBanners" runat="server" AllowFilteringByColumn="True" AllowPaging="True" AllowSorting="True" GridLines="None" Skin="Vista">
                    <GroupingSettings CaseSensitive="false" />
                    <ClientSettings>
                        <Scrolling AllowScroll="false" UseStaticHeaders="True" />
                    </ClientSettings>
                    <MasterTableView AutoGenerateColumns="False" PageSize="10" DataKeyNames="ID" CellSpacing="0" Width="80%">
                                    
                        <Columns>
                            <telerik:GridTemplateColumn HeaderText="Edit" UniqueName="TemplateColumn" AllowFiltering="false">
                                <ItemTemplate>
                                    <div align="left">
                                        <asp:ImageButton ID="BannerEditButton" ToolTip="Edit Item" OnClick="EditBannerItem" runat="server" ImageUrl="~/images/edit.gif" style="cursor:pointer;" />
                                    </div>
                                </ItemTemplate>
                                <ItemStyle Wrap="false" />
                            </telerik:GridTemplateColumn>
                            <telerik:GridBoundColumn DataField="ID" DataType="System.Int32" HeaderText="ID" ReadOnly="True" SortExpression="ID" UniqueName="ID" Visible="false">
                                <ItemStyle Wrap="False" />
                            </telerik:GridBoundColumn>
                            <telerik:GridBoundColumn DataField="Deliverable_ID" DataType="System.Int32" HeaderText="Deliverable_ID" SortExpression="Deliverable_ID" UniqueName="Deliverable_ID"  Visible="false">
                                <ItemStyle Wrap="False" />
                            </telerik:GridBoundColumn>
                            <telerik:GridBoundColumn DataField="Product_ID" DataType="System.Int32" HeaderText="Product_ID" SortExpression="Product_ID" UniqueName="Product_ID"  Visible="false">
                                <ItemStyle Wrap="False" />
                            </telerik:GridBoundColumn>
                            <telerik:GridBoundColumn DataField="Product_Instance_Name" HeaderText="Product_Instance_Name" SortExpression="Product_Instance_Name" UniqueName="Product_Instance_Name"  Visible="false">
                                <ItemStyle Wrap="False" />
                            </telerik:GridBoundColumn>
                            <telerik:GridTemplateColumn DataField="Report_Table_No" DataType="System.Decimal" HeaderText="Report_Table_No" SortExpression="Report_Table_No" UniqueName="Report_Table_No">
                                <ItemTemplate>
                                    <input type="text" style="width: 100px; font-family: tahoma; font-size: 8pt; border: 1px solid blue; text-align: right;" 
                                           onclick="ChangeBorder(this);" onfocus="ChangeBorder(this);" onblur="GoSave(this);" id='ID|<%# Eval("ID") & "|Report_Table_No|a_Banner|N|" & Eval("Report_Table_No") %>' value='<%#Eval("Report_Table_No")%>' />
                                </ItemTemplate>
                                <ItemStyle Wrap="False" />
                            </telerik:GridTemplateColumn>
                            <telerik:GridTemplateColumn DataField="Banner_Display" HeaderText="Banner_Display" SortExpression="Banner_Display" UniqueName="Banner_Display">
                                <ItemTemplate>
                                    <textarea rows="2" style="width: 450px; font-family: tahoma; font-size: 8pt; height: 25px; border: 1px solid blue; text-align: left;" 
                                              onclick="ChangeBorder(this);" onfocus="ChangeBorder(this);" onblur="GoSave(this);" id='ID|<%# Eval("ID") & "|Banner_Display|a_Banner|S|0" %>'><%#Eval("Banner_Display")%></textarea>
                                </ItemTemplate>
                                <ItemStyle Wrap="False" />
                            </telerik:GridTemplateColumn>
                            <telerik:GridTemplateColumn DataField="Order_Banner" DataType="System.Int32" HeaderText="Order_Banner" SortExpression="Order_Banner" UniqueName="Order_Banner">
                                <ItemTemplate>
                                    <input type="text" style="width: 100px; font-family: tahoma; font-size: 8pt; border: 1px solid blue; text-align: right;" 
                                           onclick="ChangeBorder(this);" onfocus="ChangeBorder(this);" onblur="GoSave(this);" id='ID|<%# Eval("ID") & "|Order_Banner|a_Banner|N|" & Eval("Order_Banner") %>' value='<%#Eval("Order_Banner")%>' />
                                </ItemTemplate>
                                <ItemStyle Wrap="False" />
                            </telerik:GridTemplateColumn>
                            <telerik:GridTemplateColumn DataField="No_of_Years" DataType="System.Decimal" HeaderText="No_of_Years" SortExpression="No_of_Years" UniqueName="No_of_Years">
                                <ItemTemplate>
                                    <input type="text" style="width: 100px; font-family: tahoma; font-size: 8pt; border: 1px solid blue; text-align: right;" 
                                           onclick="ChangeBorder(this);" onfocus="ChangeBorder(this);" onblur="GoSave(this);" id='ID|<%# Eval("ID") & "|No_of_Years|a_Banner|N|" & Eval("No_of_Years") %>' value='<%#Eval("No_of_Years")%>' />
                                </ItemTemplate>
                                <ItemStyle Wrap="False" />
                            </telerik:GridTemplateColumn>
                            <telerik:GridTemplateColumn DataField="Super_Header" HeaderText="Super_Header" SortExpression="Super_Header" UniqueName="Super_Header">
                                <ItemTemplate>
                                    <textarea rows="2" style="width: 450px; font-family: tahoma; font-size: 8pt; height: 25px; border: 1px solid blue; text-align: left;" 
                                              onclick="ChangeBorder(this);" onfocus="ChangeBorder(this);" onblur="GoSave(this);" id='ID|<%# Eval("ID") & "|Super_Header|a_Banner|S|0" %>'><%#Eval("Super_Header")%></textarea>
                                </ItemTemplate>
                                <ItemStyle Wrap="False" />
                            </telerik:GridTemplateColumn>
                            <telerik:GridTemplateColumn DataField="Source_System_Trend_id" DataType="System.Decimal" HeaderText="Source_System_Trend_id" SortExpression="Source_System_Trend_id" UniqueName="Source_System_Trend_id">
                                <ItemTemplate>
                                    <div id='ID|<%# Eval("ID") & "|Source_System_Trend_id|a_Banner|S|" & Eval("Source_System_Trend_id") %>' style="width: 100%; border: 1px solid blue;" onclick="GetTrendDropValues(this);">
                                        <select style="width: 350px; font-family: tahoma; font-size: 8pt;" id='ID|<%# Eval("ID") & "|Source_System_Trend_id|a_Banner|D|" & Eval("Source_System_Trend_id") %>' onchange="GoSave(this);">
                                            <option selected="selected" value='<%# Eval("Source_System_Trend_id") %>'><%#Eval("Source_System_Trend_id")%></option>
                                        </select>
                                    </div>
                                </ItemTemplate>
                                <ItemStyle Wrap="False" />
                            </telerik:GridTemplateColumn>
                            <telerik:GridTemplateColumn DataField="DRA_Name" HeaderText="DRA_Name" SortExpression="DRA_Name" UniqueName="DRA_Name">
                                <ItemTemplate>
                                    <div id='ID|<%# Eval("ID") & "|DRA_Name|a_Banner|S|0" %>' style="width: 450px; border: 1px solid blue;" onclick="GetDVDropValues(this);">
                                        <select style="width: 450px; font-family: tahoma; font-size: 8pt;" id='ID|<%# Eval("ID") & "|DRA_Name|a_Banner|D|0" %>' onchange="GoSaveDV(this);">
                                            <option selected="selected" value='<%# Eval("DRA_Name") %>'><%#Eval("DRA_Name")%></option>
                                        </select>
                                    </div>
                                </ItemTemplate>
                                <ItemStyle Wrap="False" />
                            </telerik:GridTemplateColumn>
                            <telerik:GridTemplateColumn DataField="DRA_Value" HeaderText="DRA_Value" SortExpression="DRA_Value" UniqueName="DRA_Value">
                                <ItemTemplate>
                                    <div id='ID|<%# Eval("ID") & "|DRA_Value|a_Banner|S|0" %>' style="width: 450px; border: 1px solid blue;" onclick="GetDVADropValues(this);">
                                        <select style="width: 450px; font-family: tahoma; font-size: 8pt;" id='ID|<%# Eval("ID") & "|DRA_Value|a_Banner|D|0" %>' onchange="GoSave(this);">
                                            <option selected="selected" value='<%# Eval("DRA_Value") %>'><%#Eval("DRA_Value")%></option>
                                        </select>
                                    </div>
                                </ItemTemplate>
                                <ItemStyle Wrap="False" />
                            </telerik:GridTemplateColumn>
                            <telerik:GridTemplateColumn DataField="Indicator" HeaderText="Indicator" SortExpression="Indicator" UniqueName="Indicator">
                                <ItemTemplate>
                                    <input type="text" style="width: 100px; font-family: tahoma; font-size: 8pt; border: 1px solid blue; text-align: right;" 
                                           onclick="ChangeBorder(this);" onfocus="ChangeBorder(this);" onblur="GoSave(this);" id='ID|<%# Eval("ID") & "|Indicator|a_Banner|S|0" %>' value='<%#Eval("Indicator")%>' />
                                </ItemTemplate>
                                <ItemStyle Wrap="False" />
                            </telerik:GridTemplateColumn>
                            <telerik:GridTemplateColumn DataField="Numerator" HeaderText="Numerator" SortExpression="Numerator" UniqueName="Numerator">
                                <ItemTemplate>
                                    <div id='ID|<%# Eval("ID") & "|Numerator|a_Banner|S|" & Eval("Numerator") %>' style="width: 100%; border: 1px solid blue;" onclick="GetDropValues(this);">
                                        <select style="width: 200px; font-family: tahoma; font-size: 8pt;" id='ID|<%# Eval("ID") & "|Numerator|a_Banner|D|" & Eval("Numerator") %>' onchange="GoSave(this);">
                                            <option selected="selected" value='<%# Eval("Numerator") %>'><%#Eval("Numerator")%></option>
                                        </select>
                                    </div>
                                </ItemTemplate>
                                <ItemStyle Wrap="False" />
                            </telerik:GridTemplateColumn>
                            <telerik:GridTemplateColumn DataField="Denominator" HeaderText="Denominator" SortExpression="Denominator" UniqueName="Denominator">
                                <ItemTemplate>
                                    <div id='ID|<%# Eval("ID") & "|Denominator|a_Banner|S|" & Eval("Denominator") %>' style="width: 100%; border: 1px solid blue;" onclick="GetDropValues(this);">
                                        <select style="width: 200px; font-family: tahoma; font-size: 8pt;" id='ID|<%# Eval("ID") & "|Denominator|a_Banner|D|" & Eval("Denominator") %>' onchange="GoSave(this);">
                                            <option selected="selected" value='<%# Eval("Denominator") %>'><%#Eval("Denominator")%></option>
                                        </select>
                                    </div>
                                </ItemTemplate>
                                <ItemStyle Wrap="False" />
                            </telerik:GridTemplateColumn>
                            <telerik:GridTemplateColumn DataField="Number_Format" HeaderText="Number_Format" SortExpression="Number_Format" UniqueName="Number_Format">
                                <ItemTemplate>
                                    <div id='ID|<%# Eval("ID") & "|Number_Format|a_Banner|S|" & Eval("Number_Format") %>' style="width: 100%; border: 1px solid blue;" onclick="GetDropValues(this);">
                                        <select style="width: 200px; font-family: tahoma; font-size: 8pt;" id='ID|<%# Eval("ID") & "|Number_Format|a_Banner|D|" & Eval("Number_Format") %>' onchange="GoSave(this);">
                                            <option selected="selected" value='<%# Eval("Number_Format") %>'><%#Eval("Number_Format")%></option>
                                        </select>
                                    </div>
                                </ItemTemplate>
                                <ItemStyle Wrap="False" />
                            </telerik:GridTemplateColumn>
                            <telerik:GridTemplateColumn DataField="Decimals" HeaderText="Decimals" SortExpression="Decimals" UniqueName="Decimals">
                                <ItemTemplate>
                                    <div id='ID|<%# Eval("ID") & "|Decimals|a_Banner|S|" & Eval("Decimals") %>' style="width: 100%; border: 1px solid blue;" onclick="GetDropValues(this);">
                                        <select style="width: 200px; font-family: tahoma; font-size: 8pt;" id='ID|<%# Eval("ID") & "|Decimals|a_Banner|D|" & Eval("Decimals") %>' onchange="GoSave(this);">
                                            <option selected="selected" value='<%# Eval("Decimals") %>'><%#Eval("Decimals")%></option>
                                        </select>
                                    </div>
                                </ItemTemplate>
                                <ItemStyle Wrap="False" />
                            </telerik:GridTemplateColumn>
                            <telerik:GridTemplateColumn DataField="Control_Column" DataType="System.Decimal" HeaderText="Control_Column" SortExpression="Control_Column" UniqueName="Control_Column">
                                <ItemTemplate>
                                    <input type="text" style="width: 100px; font-family: tahoma; font-size: 8pt; border: 1px solid blue; text-align: right;" 
                                           onclick="ChangeBorder(this);" onfocus="ChangeBorder(this);" onblur="GoSave(this);" id='ID|<%# Eval("ID") & "|Control_Column|a_Banner|N|" & Eval("Control_Column") %>' value='<%#Eval("Control_Column")%>' />
                                </ItemTemplate>
                                <ItemStyle Wrap="False" />
                            </telerik:GridTemplateColumn>
                            <telerik:GridTemplateColumn DataField="Object_Indicator" HeaderText="Object_Indicator" SortExpression="Object_Indicator" UniqueName="Object_Indicator">
                                <ItemTemplate>
                                    <input type="text" style="width: 100px; font-family: tahoma; font-size: 8pt; border: 1px solid blue; text-align: right;" 
                                           onclick="ChangeBorder(this);" onfocus="ChangeBorder(this);" onblur="GoSave(this);" id='ID|<%# Eval("ID") & "|Object_Indicator|a_Banner|S|0" %>' value='<%#Eval("Object_Indicator")%>' />
                                </ItemTemplate>
                                <ItemStyle Wrap="False" />
                            </telerik:GridTemplateColumn>
                            <telerik:GridTemplateColumn DataField="Format_Type" HeaderText="Format_Type" SortExpression="Format_Type" UniqueName="Format_Type">
                                <ItemTemplate>
                                    <div id='ID|<%# Eval("ID") & "|Format_Type|a_Banner|S|" & Eval("Format_Type") %>' style="width: 100%; border: 1px solid blue;" onclick="GetDropValues(this);">
                                        <select style="width: 200px; font-family: tahoma; font-size: 8pt;" id='ID|<%# Eval("ID") & "|Format_Type|a_Banner|N|" & Eval("Format_Type") %>' onchange="GoSave(this);">
                                            <option selected="selected" value='<%# Eval("Format_Type") %>'><%#Eval("Format_Type")%></option>
                                        </select>
                                    </div>
                                </ItemTemplate>
                                <ItemStyle Wrap="False" />
                            </telerik:GridTemplateColumn>
                            
                        </Columns>
                        <PagerStyle Mode="NextPrevNumericAndAdvanced" Position="TopAndBottom" />
                    </MasterTableView>
                    <AlternatingItemStyle BackColor="#E0E0E0" />
                </telerik:RadGrid>
            </telerik:RadPageView>
            <telerik:RadPageView ID="StubsPage" runat="server">
                <telerik:RadGrid ID="RadGridStubs" runat="server" AllowFilteringByColumn="True" AllowPaging="True" AllowSorting="True" GridLines="None" Skin="Vista">
                    <GroupingSettings CaseSensitive="false" />
                    <ClientSettings>
                        <Scrolling AllowScroll="false" UseStaticHeaders="True" />
                    </ClientSettings>
                    <MasterTableView AutoGenerateColumns="False" PageSize="10" DataKeyNames="Report_Stub_ID" CellSpacing="0" Width="80%">
                                
                        <Columns>
                            <telerik:GridTemplateColumn HeaderText="Edit" UniqueName="TemplateColumn" AllowFiltering="false">
                                <ItemTemplate>
                                    <div align="left">
                                        <asp:ImageButton ID="StubEditButton" ToolTip="Edit Item" OnClick="EditStubItem" runat="server" ImageUrl="~/images/edit.gif" style="cursor:pointer;" />
                                    </div>
                                </ItemTemplate>
                                <ItemStyle Wrap="false" />
                            </telerik:GridTemplateColumn>
                            <telerik:GridBoundColumn DataField="Report_Stub_ID" DataType="System.Int32" Visible="false" HeaderText="Report_Stub_ID" ReadOnly="True" SortExpression="Report_Stub_ID" UniqueName="Report_Stub_ID">
                                <ItemStyle Wrap="False" />
                            </telerik:GridBoundColumn>
                            <telerik:GridBoundColumn DataField="Deliverable_ID" DataType="System.Int32" Visible="false" HeaderText="Deliverable_ID" SortExpression="Deliverable_ID" UniqueName="Deliverable_ID">
                                <ItemStyle Wrap="False" />
                            </telerik:GridBoundColumn>
                            <telerik:GridTemplateColumn DataField="Stub_ID" DataType="System.Int32" Visible="true" HeaderText="Stub_ID" SortExpression="Stub_ID" UniqueName="Stub_ID">
                                <ItemTemplate>
                                    <input type="text" style="width: 100px; font-family: tahoma; font-size: 8pt; border: 1px solid blue; text-align: right;" 
                                           onclick="ChangeBorder(this);" onfocus="ChangeBorder(this);" onblur="GoSave(this);" id='Report_Stub_ID|<%# Eval("Report_Stub_ID") & "|Stub_ID|a_Stub|N|" & Eval("Stub_ID") %>' value='<%#Eval("Stub_ID")%>' />
                                </ItemTemplate>
                                <ItemStyle Wrap="False" />
                            </telerik:GridTemplateColumn>
                            <telerik:GridTemplateColumn DataField="DRA_Attr_Name" HeaderText="DRA_Attr_Name" SortExpression="DRA_Attr_Name" UniqueName="DRA_Attr_Name">
                                <ItemTemplate>
                                    <div id='Report_Stub_ID|<%# Eval("Report_Stub_ID") & "|DRA_Attr_Name|a_Stub|S|0" %>' style="width: 450px; border: 1px solid blue;" onclick="GetDVDropValues(this);">
                                        <select style="width: 450px; font-family: tahoma; font-size: 8pt;" id='Report_Stub_ID|<%# Eval("Report_Stub_ID") & "|DRA_Attr_Name|a_Stub|D|0" %>' onchange="GoSaveDV(this);">
                                            <option selected="selected" value='<%# Eval("DRA_Attr_Name") %>'><%#Eval("DRA_Attr_Name")%></option>
                                        </select>
                                    </div>
                                </ItemTemplate>
                                <ItemStyle Wrap="False" />
                            </telerik:GridTemplateColumn>
                            <telerik:GridTemplateColumn DataField="DRA_Attr_Value" HeaderText="DRA_Attr_Value" SortExpression="DRA_Attr_Value" UniqueName="DRA_Attr_Value">
                                <ItemTemplate>
                                    <div id='Report_Stub_ID|<%# Eval("Report_Stub_ID") & "|DRA_Attr_Value|a_Stub|S|0" %>' style="width: 450px; border: 1px solid blue;" onclick="GetDVADropValues(this);">
                                        <select style="width: 450px; font-family: tahoma; font-size: 8pt;" id='Report_Stub_ID|<%# Eval("Report_Stub_ID") & "|DRA_Attr_Value|a_Stub|D|0" %>' onchange="GoSave(this);">
                                            <option selected="selected" value='<%# Eval("DRA_Attr_Value") %>'><%#Eval("DRA_Attr_Value")%></option>
                                        </select>
                                    </div>
                                </ItemTemplate>
                                <ItemStyle Wrap="False" />
                            </telerik:GridTemplateColumn>
                            <telerik:GridTemplateColumn DataField="Sort_Attr_Name" HeaderText="Sort_Attr_Name" SortExpression="Sort_Attr_Name" UniqueName="Sort_Attr_Name" DataType="System.Int32">
                                <ItemTemplate>
                                    <input type="text" style="width: 100px; font-family: tahoma; font-size: 8pt; border: 1px solid blue; text-align: right;" 
                                           onclick="ChangeBorder(this);" onfocus="ChangeBorder(this);" onblur="GoSave(this);" id='Report_Stub_ID|<%# Eval("Report_Stub_ID") & "|Sort_Attr_Name|a_Stub|N|" & Eval("Sort_Attr_Name") %>' value='<%#Eval("Sort_Attr_Name")%>' />
                                </ItemTemplate>
                                <ItemStyle Wrap="False" />
                            </telerik:GridTemplateColumn>
                            <telerik:GridTemplateColumn DataField="Sort_Attr_Value" DataType="System.Int32" HeaderText="Sort_Attr_Value" SortExpression="Sort_Attr_Value" UniqueName="Sort_Attr_Value">
                                <ItemTemplate>
                                    <input type="text" style="width: 100px; font-family: tahoma; font-size: 8pt; border: 1px solid blue; text-align: right;" 
                                           onclick="ChangeBorder(this);" onfocus="ChangeBorder(this);" onblur="GoSave(this);" id='Report_Stub_ID|<%# Eval("Report_Stub_ID") & "|Sort_Attr_Value|a_Stub|N|" & Eval("Sort_Attr_Value") %>' value='<%#Eval("Sort_Attr_Value")%>' />
                                </ItemTemplate>
                                <ItemStyle Wrap="False" />
                            </telerik:GridTemplateColumn>
                            <telerik:GridTemplateColumn DataField="Attr_Name_Override" HeaderText="Attr_Name_Override" SortExpression="Attr_Name_Override" UniqueName="Attr_Name_Override">
                                <ItemTemplate>
                                    <textarea rows="2" style="width: 450px; font-family: tahoma; font-size: 8pt; height: 25px; border: 1px solid blue; text-align: left;" 
                                              onclick="ChangeBorder(this);" onfocus="ChangeBorder(this);" onblur="GoSave(this);" id='Report_Stub_ID|<%# Eval("Report_Stub_ID") & "|Attr_Name_Override|a_Stub|S|0" %>'><%#Eval("Attr_Name_Override")%></textarea>
                                </ItemTemplate>
                                <ItemStyle Wrap="False" />
                            </telerik:GridTemplateColumn>
                            <telerik:GridTemplateColumn DataField="Attr_Value_Override" HeaderText="Attr_Value_Override" SortExpression="Attr_Value_Override" UniqueName="Attr_Value_Override">
                                <ItemTemplate>
                                    <textarea rows="2" style="width: 450px; font-family: tahoma; font-size: 8pt; height: 25px; border: 1px solid blue; text-align: left;" 
                                              onclick="ChangeBorder(this);" onfocus="ChangeBorder(this);" onblur="GoSave(this);" id='Report_Stub_ID|<%# Eval("Report_Stub_ID") & "|Attr_Value_Override|a_Stub|S|0" %>'><%# Eval("Attr_Value_Override")%></textarea>
                                </ItemTemplate>
                                <ItemStyle Wrap="False" />
                            </telerik:GridTemplateColumn>
                            <telerik:GridTemplateColumn DataField="Super_Header" HeaderText="Super_Header" SortExpression="Super_Header" UniqueName="Super_Header">
                                <ItemTemplate>
                                    <textarea rows="2" style="width: 450px; font-family: tahoma; font-size: 8pt; height: 25px; border: 1px solid blue; text-align: left;" 
                                              onclick="ChangeBorder(this);" onfocus="ChangeBorder(this);" onblur="GoSave(this);" id='Report_Stub_ID|<%# Eval("Report_Stub_ID") & "|Super_Header|a_Stub|S|0" %>'><%#Eval("Super_Header")%></textarea>
                                </ItemTemplate>
                                <ItemStyle Wrap="False" />
                            </telerik:GridTemplateColumn>
                            <telerik:GridTemplateColumn DataField="Name_Value_Flag" HeaderText="Name_Value_Flag" SortExpression="Name_Value_Flag" UniqueName="Name_Value_Flag">
                                <ItemTemplate>
                                    <input type="text" style="width: 100px; font-family: tahoma; font-size: 8pt; border: 1px solid blue; text-align: right;" 
                                           onclick="ChangeBorder(this);" onfocus="ChangeBorder(this);" onblur="GoSave(this);" id='Report_Stub_ID|<%# Eval("Report_Stub_ID") & "|Name_Value_Flag|a_Stub|S|0" %>' value='<%#Eval("Name_Value_Flag")%>' />
                                </ItemTemplate>
                                <ItemStyle Wrap="False" />
                            </telerik:GridTemplateColumn>
                            <telerik:GridTemplateColumn DataField="Control_Column" HeaderText="Control_Column" SortExpression="Control_Column" UniqueName="Control_Column" DataType="System.Decimal">
                                <ItemTemplate>
                                    <input type="text" style="width: 100px; font-family: tahoma; font-size: 8pt; border: 1px solid blue; text-align: right;" 
                                           onclick="ChangeBorder(this);" onfocus="ChangeBorder(this);" onblur="GoSave(this);" id='Report_Stub_ID|<%# Eval("Report_Stub_ID") & "|Control_Column|a_Stub|N|" & Eval("Control_Column") %>' value='<%#Eval("Control_Column")%>' />
                                </ItemTemplate>
                                <ItemStyle Wrap="False" />
                            </telerik:GridTemplateColumn>
                            <telerik:GridTemplateColumn DataField="Format_Type" HeaderText="Format_Type" SortExpression="Format_Type" UniqueName="Format_Type">
                                <ItemTemplate>
                                    <div id='Report_Stub_ID|<%# Eval("Report_Stub_ID") & "|Format_Type|a_stub|S|" & Eval("Format_Type") %>' style="width: 100%; border: 1px solid blue;" onclick="GetDropValues(this);">
                                        <select style="width: 200px; font-family: tahoma; font-size: 8pt;" id='ID|<%# Eval("Report_Stub_ID") & "|Format_Type|a_stub|N|" & Eval("Format_Type") %>' onchange="GoSave(this);">
                                            <option selected="selected" value='<%# Eval("Format_Type") %>'><%#Eval("Format_Type")%></option>
                                        </select>
                                    </div>
                                </ItemTemplate>
                                <ItemStyle Wrap="False" />
                            </telerik:GridTemplateColumn>
                        </Columns>
                        <PagerStyle Mode="NextPrevNumericAndAdvanced" Position="TopAndBottom" />
                    </MasterTableView>
                    <AlternatingItemStyle BackColor="#E0E0E0" />
                </telerik:RadGrid>
            </telerik:RadPageView>
            <telerik:RadPageView ID="pageImportExport" runat="server">
                <div style="text-align: center;">
                    <asp:Label runat="server" ID="lblError" Font-Bold="true" ForeColor="Red" Font-Names="Arial" Font-Size="10pt"></asp:Label>
                    <fieldset name="Group1" style="width: 400px; font-family: Arial; font-size: 10pt;">
                        <legend>
                            Export: <strong>
                                <asp:Literal runat="server" ID="legendText"></asp:Literal>
                            </strong>
                        </legend>
                                
                        <table style="width: 300px">
                            <tr>
                                <td style="text-align: right; width: 150px;">Tables</td>
                                <td style="text-align: left; width: 150px;">
                                    <asp:CheckBox id="chkReports" runat="server" Checked="True" />
                                </td>
                            </tr>
                            <tr>
                                <td style="text-align: right; width: 150px;">Banners</td>
                                <td style="text-align: left; width: 150px;">
                                    <asp:CheckBox id="chkBanners" runat="server" Checked="True" />
                                </td>
                            </tr>
                            <tr>
                                <td style="text-align: right; width: 150px;">Stubs</td>
                                <td style="text-align: left; width: 150px;">
                                    <asp:CheckBox id="chkStubs" runat="server" Checked="True" />
                                </td>
                            </tr>
                            <tr>
                                <td colspan="2" style="width: 300px;">
                                    <asp:Button id="btnExport" runat="server" Text="Export" />
                                </td>
                            </tr>
                        </table><br />
                    </fieldset>
                </div><br />
            </telerik:RadPageView>
            <telerik:RadPageView ID="NewReport" runat="server">
                <table style="font-family: 'Segoe UI'; font-size: 8pt; color: #000000; text-align: right; vertical-align: top; font-weight: normal; white-space: nowrap">
                    <tr>
                        <td>
                            <strong>Table ID:</strong>
                        </td>
                        <td style="text-align: left;">
                            <asp:Label id="lblTableID" runat="server" Text=""></asp:Label>
                        </td>
                    </tr>
                    <tr>
                        <td>
                            <strong>Deliverable ID:</strong>
                        </td>
                        <td style="text-align: left;">
                            <asp:Label id="lblDeliverableID" runat="server" Text=""></asp:Label>
                        </td>
                    </tr>
                    <tr>
                        <td>
                            <strong>Product ID:</strong>
                        </td>
                        <td style="text-align: left;">
                            <asp:Label id="lblProductID" runat="server" Text=""></asp:Label>
                        </td>
                    </tr>
                    <tr>
                        <td>
                            <strong>Product Instance Name:</strong>
                        </td>
                        <td style="text-align: left;">
                            <asp:Label id="lblProductInstanceName" runat="server" Text=""></asp:Label>
                        </td>
                    </tr>
                    <tr>
                        <td>
                            <strong>Top N:</strong>
                        </td>
                        <td style="text-align: left;">
                            <telerik:RadNumericTextBox ID="txtTopN" runat="server" MaxValue="99999" MinValue="-99999" Skin="Web20" Height="12px" Width="100px">
                                <NumberFormat AllowRounding="False" DecimalDigits="2" GroupSeparator="" GroupSizes="6" KeepNotRoundedValue="True" NegativePattern="-n" PositivePattern="n" />
                            </telerik:RadNumericTextBox>
                        </td>
                    </tr>
                    <tr>
                        <td>
                            <strong>Report Table No:</strong>
                        </td>
                        <td style="text-align: left;">
                            <telerik:RadNumericTextBox ID="txtReportTableNo" runat="server" MaxValue="99999" Height="12px" MinValue="-99999" Skin="Web20" Width="100px">
                                <NumberFormat AllowRounding="False" DecimalDigits="0" GroupSeparator="" GroupSizes="6" />
                            </telerik:RadNumericTextBox>
                        </td>
                    </tr>
                    <tr>
                        <td>
                            <strong>Report Name:</strong>
                        </td>
                        <td style="text-align: left;">
                            <asp:TextBox id="txtReportName" runat="server" Font-Names="Segoe UI" Font-Size="8pt" Height="30px" TextMode="MultiLine" Width="550px"></asp:TextBox>
                        </td>
                    </tr>
                    <tr>
                        <td>
                            <strong>Report Footer:</strong>
                        </td>
                        <td style="text-align: left;">
                            <asp:TextBox id="txtReportFooter1" runat="server" Font-Names="Segoe UI" Font-Size="8pt" Height="30px" TextMode="MultiLine" MaxLength="1500" Width="550px"></asp:TextBox>
                        </td>
                    </tr>
                    <tr>
                        <td>
                            <strong>Report Footer 2:</strong>
                        </td>
                        <td style="text-align: left;">
                            <asp:TextBox id="txtReportFooter2" runat="server" Font-Names="Segoe UI" Font-Size="8pt" Height="30px" TextMode="MultiLine" Width="550px"></asp:TextBox>
                        </td>
                    </tr>
                    <tr>
                        <td>
                            <strong>Report Footer 3:</strong>
                        </td>
                        <td style="text-align: left;">
                            <asp:TextBox id="txtReportFooter3" runat="server" Font-Names="Segoe UI" Font-Size="8pt" Height="30px" TextMode="MultiLine" Width="550px"></asp:TextBox>
                        </td>
                    </tr>
                    <tr>
                        <td>
                            <strong>Market Segment Name:</strong>
                        </td>
                        <td style="text-align: left;">
                            <asp:DropDownList runat="server" ID="MKTSEG_Name" Width="450px">
                            </asp:DropDownList>
                        </td>
                    </tr>
                    <tr>
                        <td>
                            <strong>Market Segment Value:</strong>
                        </td>
                        <td style="text-align: left;">
                            <div runat="server" name="MSValueHolder" ID="MSValueHolder"></div>
                        
                        </td>
                    </tr>
                    <tr>
                        <td>
                            <strong>Market Segment Display:</strong>
                        </td>
                        <td style="text-align: left;">
                            <asp:TextBox id="txtMSDisplay" runat="server" Font-Names="Segoe UI" Font-Size="8pt" Height="30px" TextMode="MultiLine" Width="550px"></asp:TextBox>
                        </td>
                    </tr>
                    <tr>
                        <td>
                            <strong>Rank Type:</strong>
                        </td>
                        <td style="text-align: left;">
                            <asp:TextBox ID="txtRankType" runat="server" Font-Names="Segoe UI" Font-Size="8pt" MaxLength="50" Width="200px"></asp:TextBox>
                        </td>
                    </tr>
                    <tr>
                        <td>
                            <strong>Sort Year:</strong>
                        </td>
                        <td style="text-align: left;">
                            <telerik:RadNumericTextBox ID="txtSortYear" runat="server" MaxValue="9999" Height="12px" MinValue="-9999" Skin="Web20" Width="100px">
                                <NumberFormat AllowRounding="False" DecimalDigits="0" GroupSeparator="" GroupSizes="6" />
                            </telerik:RadNumericTextBox>
                        </td>
                    </tr>
                    <tr>
                        <td>
                            <strong>Stub ID:</strong>
                        </td>
                        <td style="text-align: left;">
                            <telerik:RadNumericTextBox ID="txtStubID" runat="server" MaxValue="999999999999" Height="12px" MinValue="-999999999999" Skin="Web20" Width="100px">
                                <NumberFormat AllowRounding="False" DecimalDigits="0" GroupSeparator="" GroupSizes="9" />
                            </telerik:RadNumericTextBox>
                        </td>
                    </tr>
                    <tr>
                        <td>
                            <strong>Table Type No:</strong>
                        </td>
                        <td style="text-align: left;">
                            <asp:DropDownList ID="Table_Type_No" runat="server"></asp:DropDownList>
                        </td>
                    </tr>
                    <tr>
                        <td>
                            <strong>Table Average 1:</strong>
                        </td>
                        <td style="text-align: left;">
                            <asp:DropDownList ID="Table_Average_1" runat="server"></asp:DropDownList>
                        </td>
                    </tr>
                    <tr>
                        <td>
                            <strong>Table Average 2:</strong>
                        </td>
                        <td style="text-align: left;">
                            <asp:DropDownList runat="server" ID="Table_Average_2"></asp:DropDownList>
                        </td>
                    </tr>
                    <tr>
                        <td>
                            <strong>Table Average 3:</strong>
                        </td>
                        <td style="text-align: left;">
                            <asp:DropDownList runat="server" ID="Table_Average_3"></asp:DropDownList>
                        </td>
                    </tr>
                    <tr>
                        <td>
                            <strong>Table Banner Height:</strong>
                        </td>
                        <td style="text-align: left;">
                            <telerik:RadNumericTextBox ID="txtTableBannerHeight" runat="server" MaxValue="99999" Height="12px" MinValue="-99999" Skin="Web20" Width="100px">
                                <NumberFormat AllowRounding="False" DecimalDigits="0" GroupSeparator="" GroupSizes="6" />
                            </telerik:RadNumericTextBox>
                        </td>
                    </tr>
                    <tr>
                        <td>
                            <strong>Unique ID:</strong>
                        </td>
                        <td style="text-align: left;">
                            <asp:TextBox id="txtUniqueID" runat="server" Font-Names="Segoe UI" Font-Size="8pt" Width="250px"></asp:TextBox>
                        </td>
                    </tr>
                    <tr>
                        <td>
                            <strong>Group ID:</strong>
                        </td>
                        <td style="text-align: left;">
                            <asp:TextBox id="txtGroupID" runat="server" Font-Names="Segoe UI" Font-Size="8pt" Width="250px"></asp:TextBox>
                        </td>
                    </tr>
                    <tr>
                        <td>
                            <strong>Group Name:</strong>
                        </td>
                        <td style="text-align: left;">
                            <asp:TextBox id="txtGroupName" runat="server" Font-Names="Segoe UI" Font-Size="8pt" Width="250px"></asp:TextBox>
                        </td>
                    </tr>
                    <tr>
                        <td style="text-align: center;" colspan="2">
                            <asp:Button id="btnRSave" runat="server" Text="Save" Font-Names="segoi ui" Font-Size="10pt" Height="25px" Width="59px" />
                            &nbsp;&nbsp; &nbsp;
                            <asp:Button id="btnRDelete" runat="server" Text="Delete" Font-Names="segoi ui" Font-Size="10pt" Height="25px" Width="59px" />
                            &nbsp; &nbsp;
                            <asp:Button id="btnRCancel" runat="server" Text="Cancel" Font-Names="segoi ui" Font-Size="10pt" Height="25px" Width="59px" />
                        </td>
                    </tr>
                </table>
                        
            </telerik:RadPageView>
            <telerik:RadPageView ID="NewBanner" runat="server">
                <table style="font-family: 'Segoe UI'; font-size: 8pt; color: #000000; text-align: right; vertical-align: top; font-weight: normal; white-space: nowrap">
                    <tr>
                        <td>
                            <strong>Banner ID:</strong>
                        </td>
                        <td style="text-align: left;">
                            <asp:Label id="lblBannerID" runat="server" Text=""></asp:Label>
                        </td>
                    </tr>
                    <tr>
                        <td>
                            <strong>Deliverable ID:</strong>
                        </td>
                        <td style="text-align: left;">
                            <asp:Label id="lblBannerDeliverableID" runat="server" Text=""></asp:Label>
                        </td>
                    </tr>
                    <tr>
                        <td>
                            <strong>Product ID:</strong>
                        </td>
                        <td style="text-align: left;">
                            <asp:Label id="lblBannerProductID" runat="server" Text=""></asp:Label>
                        </td>
                    </tr>
                    <tr>
                        <td>
                            <strong>Product Instance Name:</strong>
                        </td>
                        <td style="text-align: left;">
                            <asp:Label id="lblBannerProductInstanceName" runat="server" Text=""></asp:Label>
                            
                        </td>
                    </tr>
                    <tr>
                        <td>
                            <strong>Report Table No:</strong>
                        </td>
                        <td style="text-align: left;">
                            <asp:DropDownList runat="server" ID="banner_report_table_no" Width="450px">
                            </asp:DropDownList>
                        </td>
                    </tr>
                    <tr>
                        <td>
                            <strong>Banner Display:</strong>
                        </td>
                        <td style="text-align: left;">
                            <asp:TextBox id="txtBannerDisplay" runat="server" Font-Names="Segoe UI" Font-Size="8pt" Height="30px" TextMode="MultiLine" Width="550px"></asp:TextBox>
                        </td>
                    </tr>
                    
                    <tr>
                        <td>
                            <strong>Order Banner:</strong>
                        </td>
                        <td style="text-align: left;">
                            <telerik:RadNumericTextBox ID="txtOrderBanner" runat="server" MaxValue="9999999999" Height="12px" MinValue="-9999999999" Skin="Web20" Width="100px">
                                <NumberFormat AllowRounding="False" DecimalDigits="0" GroupSeparator="" GroupSizes="6" />
                            </telerik:RadNumericTextBox>
                        </td>
                    </tr>
                    <tr>
                        <td>
                            <strong>Number Of Years:</strong>
                        </td>
                        <td style="text-align: left;">
                            <telerik:RadNumericTextBox ID="txtNumOfYears" runat="server" MaxValue="9999" Height="12px" MinValue="-9999" Skin="Web20" Width="100px">
                                <NumberFormat AllowRounding="False" DecimalDigits="0" GroupSeparator="" GroupSizes="6" />
                            </telerik:RadNumericTextBox>
                        </td>
                    </tr>
                    <tr>
                        <td>
                            <strong>Super Header:</strong>
                        </td>
                        <td style="text-align: left;">
                            <asp:TextBox id="txtSuperHeader" runat="server" Font-Names="Segoe UI" Font-Size="8pt" Height="30px" TextMode="MultiLine" Width="550px"></asp:TextBox>
                        </td>
                    </tr>
                    <tr>
                        <td>
                            <strong>Source System Trend ID:</strong>
                        </td>
                        <td style="text-align: left;">
                            <asp:DropDownList runat="server" ID="Source_System_Trend_ID" Width="450px">
                            </asp:DropDownList>
                        </td>
                    </tr>
                    <tr>
                        <td>
                            <strong>DRA Name:</strong>
                        </td>
                        <td style="text-align: left;">
                            <asp:DropDownList runat="server" ID="DRA_Name" Width="450px">
                            </asp:DropDownList>
                        </td>
                    </tr>
                    <tr>
                        <td>
                            <strong>DRA Value:</strong>
                        </td>
                        <td style="text-align: left;">
                            <div runat="server" name="DRAValueHolder" ID="DRAValueHolder"></div>
                        </td>
                    </tr>
                    <tr>
                        <td>
                            <strong>Indicator:</strong>
                        </td>
                        <td style="text-align: left;">
                            <asp:TextBox id="txtIndicator" runat="server" Font-Names="Segoe UI" Font-Size="8pt" MaxLength="20" Width="200px"></asp:TextBox>
                        </td>
                    </tr>
                    <tr>
                        <td>
                            <strong>Numerator:</strong>
                        </td>
                        <td style="text-align: left;">
                            <asp:DropDownList ID="Numerator" runat="server"></asp:DropDownList>
                        </td>
                    </tr>
                    <tr>
                        <td>
                            <strong>Denominator:</strong>
                        </td>
                        <td style="text-align: left;">
                            <asp:DropDownList ID="Denominator" runat="server"></asp:DropDownList>
                        </td>
                    </tr>
                    <tr>
                        <td>
                            <strong>Number Format:</strong>
                        </td>
                        <td style="text-align: left;">
                            <asp:DropDownList ID="Number_Format" runat="server"></asp:DropDownList>
                        </td>
                    </tr>
                    <tr>
                        <td>
                            <strong>Decimals:</strong>
                        </td>
                        <td style="text-align: left;">
                            <asp:DropDownList ID="Decimals" runat="server"></asp:DropDownList>
                        </td>
                    </tr>
                    <tr>
                        <td>
                            <strong>Control Column:</strong>
                        </td>
                        <td style="text-align: left;">
                            <telerik:RadNumericTextBox ID="txtControlColumn" runat="server" MaxValue="9999999999" Height="12px" MinValue="-9999999999" Skin="Web20" Width="100px">
                                <NumberFormat AllowRounding="False" DecimalDigits="0" GroupSeparator="" GroupSizes="6" />
                            </telerik:RadNumericTextBox>
                        </td>
                    </tr>
                    <tr>
                        <td>
                            <strong>Object Indicator:</strong>
                        </td>
                        <td style="text-align: left;">
                            <asp:TextBox id="txtObjectIndicator" runat="server" Font-Names="Segoe UI" Font-Size="8pt" MaxLength="20" Width="200px"></asp:TextBox>
                        </td>
                    </tr>
                    <tr>
                        <td>
                            <strong>Format Type:</strong>
                        </td>
                        <td style="text-align: left;">
                            <asp:DropDownList ID="lstFormatType" runat="server"></asp:DropDownList>
                        </td>
                    </tr>
                    <tr>
                        <td style="text-align: center;" colspan="2">
                            <asp:Button id="btnBannerSave" runat="server" Text="Save" Font-Names="segoi ui" Font-Size="10pt" Height="25px" Width="59px" />
                            &nbsp;&nbsp; &nbsp;
                            <asp:Button id="btnBannerDelete" runat="server" Text="Delete" Font-Names="segoi ui" Font-Size="10pt" Height="25px" Width="59px" />
                            &nbsp; &nbsp;
                            <asp:Button id="btnBannerCancel" runat="server" Text="Cancel" Font-Names="segoi ui" Font-Size="10pt" Height="25px" Width="59px" />
                        </td>
                    </tr>
                </table>
            </telerik:RadPageView>
            <telerik:RadPageView ID="NewStub" runat="server">
                <table style="font-family: 'Segoe UI'; font-size: 8pt; color: #000000; text-align: right; vertical-align: top; font-weight: normal; white-space: nowrap">
                    <tr>
                        <td>
                            <strong>Report Stub ID:</strong>
                        </td>
                        <td style="text-align: left;">
                            <asp:Label id="lblReportStubID" runat="server" Text="Label"></asp:Label>
                        </td>
                    </tr>
                    <tr>
                        <td>
                            <strong>Deliverable ID:</strong>
                        </td>
                        <td style="text-align: left;">
                            <asp:Label id="lblStubDeliverableID" runat="server" Text="Label"></asp:Label>
                        </td>
                    </tr>
                    <tr>
                        <td>
                            <strong>Stub ID:</strong>
                        </td>
                        <td style="text-align: left;">
                            <telerik:RadNumericTextBox ID="Stub_Stub_ID" runat="server" MaxValue="9999999999" Height="12px" MinValue="-9999999999" Skin="Web20" Width="100px">
                                <NumberFormat AllowRounding="False" DecimalDigits="0" GroupSeparator="" GroupSizes="6" />
                            </telerik:RadNumericTextBox>
                        </td>
                    </tr>
                    <tr>
                        <td>
                            <strong>DRA Attr Name:</strong>
                        </td>
                        <td style="text-align: left;">
                            <asp:DropDownList runat="server" ID="dra_attr_name" Width="450px">
                            </asp:DropDownList>
                        </td>
                    </tr>
                    <tr>
                        <td>
                            <strong>DRA Attr Value:</strong>
                        </td>
                        <td style="text-align: left;">
                            <div runat="server" name="DRAAttrValueHolder" ID="DRAAttrValueHolder"></div>
                        </td>
                    </tr>
                    <tr>
                        <td>
                            <strong>Sort Attr Name:</strong>
                        </td>
                        <td style="text-align: left;">
                            <telerik:RadNumericTextBox ID="txtSortAttrName" runat="server" MaxValue="9999999999" Height="12px" MinValue="-9999999999" Skin="Web20" Width="100px">
                                <NumberFormat AllowRounding="False" DecimalDigits="0" GroupSeparator="" GroupSizes="6" />
                            </telerik:RadNumericTextBox>
                        </td>
                    </tr>
                        
                    <tr>
                        <td>
                            <strong>Sort Attr Value:</strong>
                        </td>
                        <td style="text-align: left;">
                            <telerik:RadNumericTextBox ID="txtSortAttrValue" runat="server" MaxValue="9999999999" Height="12px" MinValue="-9999999999" Skin="Web20" Width="100px">
                                <NumberFormat AllowRounding="False" DecimalDigits="0" GroupSeparator="" GroupSizes="6" />
                            </telerik:RadNumericTextBox>
                        </td>
                    </tr>
                    <tr>
                        <td>
                            <strong>Attr Name Override:</strong>
                        </td>
                        <td style="text-align: left;">
                            <asp:TextBox id="txtAttrNameOverride" runat="server" Font-Names="Segoe UI" Font-Size="8pt" Height="30px" TextMode="MultiLine" Width="550px"></asp:TextBox>

                        </td>
                    </tr>
                    <tr>
                        <td>
                            <strong>Attr Value Override:</strong>
                        </td>
                        <td style="text-align: left;">
                            <asp:TextBox id="txtAttr_Value_Override" runat="server" Font-Names="Segoe UI" Font-Size="8pt" Height="30px" TextMode="MultiLine" Width="550px"></asp:TextBox>
                        
                        </td>
                    </tr>
                    <tr>
                        <td>
                            <strong>Super Header:</strong>
                        </td>
                        <td style="text-align: left;">
                            <asp:TextBox id="Stub_Super_Header" runat="server" Font-Names="Segoe UI" Font-Size="8pt" Height="30px" TextMode="MultiLine" Width="550px"></asp:TextBox>
                        </td>
                    </tr>
                    <tr>
                        <td>
                            <strong>Name Value Flag:</strong>
                        </td>
                        <td style="text-align: left;">
                            <asp:TextBox id="txtNameValueFlag" runat="server" Font-Names="Segoe UI" Font-Size="8pt" MaxLength="20" Width="200px"></asp:TextBox>
                        </td>
                    </tr>
                    <tr>
                        <td>
                            <strong>Control Column:</strong>
                        </td>
                        <td style="text-align: left;">
                            <telerik:RadNumericTextBox ID="Stub_Control_Number" runat="server" MaxValue="9999999999" Height="12px" MinValue="-9999999999" Skin="Web20" Width="100px">
                                <NumberFormat AllowRounding="False" DecimalDigits="0" GroupSeparator="" GroupSizes="6" />
                            </telerik:RadNumericTextBox>
                        </td>
                    </tr>
                    <tr>
                        <td>
                            <strong>Format Type:</strong>
                        </td>
                        <td style="text-align: left;">
                            <asp:DropDownList ID="lstFormatType2" runat="server"></asp:DropDownList>
                        </td>
                    </tr>
                    <tr>
                        <td style="text-align: center;" colspan="2">
                            <asp:Button id="btnStubSave" runat="server" Text="Save" Font-Names="segoi ui" Font-Size="10pt" Height="25px" Width="59px" />
                            &nbsp;&nbsp; &nbsp;
                            <asp:Button id="btnStubDelete" runat="server" Text="Delete" Font-Names="segoi ui" Font-Size="10pt" Height="25px" Width="59px" />
                            &nbsp; &nbsp;
                            <asp:Button id="btnStubCancel" runat="server" Text="Cancel" Font-Names="segoi ui" Font-Size="10pt" Height="25px" Width="59px" />
                        </td>
                    </tr>
                </table>
            </telerik:RadPageView>
            <telerik:RadPageView ID="ResortReports" runat="server">
                <telerik:radtreeview runat="server" ID="radTreeResort"  EnableDragAndDrop="true" Height="85%" ShowLineImages="false"
                                     MultipleSelect="false" EnableDragAndDropBetweenNodes="true"  OnClientNodeDropping="ClientNodeDropping">
                </telerik:radtreeview>
                <div align="center">
                    <asp:Button id="btnSortReports" runat="server" Text="Re-Sort" Font-Names="segoi ui" Font-Size="10pt" Height="25px" Width="59px" />
                    &nbsp;&nbsp; &nbsp;
                    <asp:Button id="btnResortCancel" runat="server" Text="Cancel / Return to Reports Grid" Font-Names="segoi ui" Font-Size="10pt" Height="25px" Width="160px" /><br />
                </div>
            </telerik:RadPageView>
        </telerik:RadMultiPage>
    </telerik:RadAjaxPanel>
</asp:Content>

                    