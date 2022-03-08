<%@ Page Title="Home Page" Language="C#" MasterPageFile="~/Site.Master" AutoEventWireup="true" CodeBehind="Default.aspx.cs" Inherits="ExportSample._Default" %>

<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">

    <div class="row">
        <div class="col-md-12">
            <h3>Select option to export the data:</h3>
        </div>
    </div>

    <div class="row">
        <div class="col-md-12">
            <div class="buttonBar">
                <asp:DropDownList runat="server" ID="ddlReportType" ClientIDMode="Static">
                    <asp:ListItem Value="-1">- Select -</asp:ListItem>
                    <asp:ListItem Value="1">Excel Report</asp:ListItem>
                    <asp:ListItem Value="2">CSV Report</asp:ListItem>
                    <asp:ListItem Value="3">Text Report</asp:ListItem>
                    <asp:ListItem Value="4">PDF Report</asp:ListItem>
                    <asp:ListItem Value="5">HTML Report</asp:ListItem>
                </asp:DropDownList>
                <asp:Button runat="server" ID="btnExportReport" Text="Export" OnClick="BtnExportReportClick"
                    ValidationGroup="export" CssClass="btn btn-primary btn-sm" />
                <asp:RequiredFieldValidator ID="reqReportType" runat="server" Text="* Please select report type to export"
                    ValidationGroup="export" Display="Dynamic" ControlToValidate="ddlReportType"
                    ForeColor="Red" InitialValue="-1"></asp:RequiredFieldValidator>
            </div>
        </div>
    </div>

</asp:Content>
