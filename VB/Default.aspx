<%@ Page Language="vb" AutoEventWireup="true" CodeFile="Default.aspx.vb" Inherits="_Default" %>

<%@ Register Assembly="DevExpress.Web.ASPxSpreadsheet.v16.1, Version=16.1.17.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" Namespace="DevExpress.Web.ASPxSpreadsheet" TagPrefix="dx" %>

<%@ Register Assembly="DevExpress.Web.v16.1, Version=16.1.17.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" Namespace="DevExpress.Web" TagPrefix="dx" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
    <script type="text/javascript">
        function onSelectedIndexChanged(s,e){
            spreadsheet.PerformCallback(s.GetValue());
        }

    </script>

</head>
<body>
    <form id="form1" runat="server">
            <dx:ASPxRadioButtonList ID="RadioButtonList" runat="server" ValueType="System.String" CssClass="myList" RepeatColumns="3" >
               <ClientSideEvents SelectedIndexChanged="onSelectedIndexChanged" />
                 <Items>
                    <dx:ListEditItem Text="Pie Chart" Value="PieChart" />
                    <dx:ListEditItem Text="Bar Chart" Value="BarChart" />
                    <dx:ListEditItem Text="Column Chart" Value="ColumnChart" />
                    <dx:ListEditItem Text="Complex Chart" Value="ComplexChart" />
                    <dx:ListEditItem Text="Doughnut Chart" Value="DoughnutChart" />
                    <dx:ListEditItem Text="Pie3d Chart" Value="Pie3dChart" />
                    <dx:ListEditItem Text="Scatter Chart" Value="ScatterChart" />
                    <dx:ListEditItem Text="Stock Chart" Value="StockChart" />
                    <dx:ListEditItem Text="Bubble Chart" Value="BubbleChart" />
                </Items>
            </dx:ASPxRadioButtonList>
        <br />
            <dx:ASPxSpreadsheet ID="Spreadsheet" ClientInstanceName="spreadsheet" runat="server" WorkDirectory="~/App_Data/WorkDirectory" Height="600px" Width="1100px" OnCallback="Spreadsheet_Callback">
            </dx:ASPxSpreadsheet>

    </form>
</body>
</html>