<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="ExcelUpload.aspx.cs" Inherits="ExcelFilter.ExcelUpload" MasterPageFile="~/Site.Master"%>

<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">

    <div class="row">
        <h3>Upload Excel File</h3>
        <br />
        <asp:FileUpload ID="flp_Upload" runat="server"/>
        <br />
        <asp:Button ID="btn_Upload" runat="server" Text="Upload Excel File" OnClick="btn_Upload_Click"/>
        <br />
        <hr />
        <asp:Label ID="lbl_Sheet" runat="server" Text="Choose Sheet"></asp:Label>
        <asp:DropDownList ID="drp_Sheet" runat="server" OnSelectedIndexChanged="drp_Sheet_SelectedIndexChanged" AutoPostBack="true"></asp:DropDownList>
        <br /><br />
        <asp:Label ID="lbl_Colums" runat="server" Text="Choose Column"></asp:Label>
        <asp:DropDownList ID="drp_Column" runat="server"></asp:DropDownList>
        <hr />
        <asp:Button ID="btn_Retrieve" runat="server" Text="Retrieve Data" OnClick="btn_Retrieve_Click"/>
        <hr />
        <h3>Retrieved Results</h3>
        <br />
        <asp:GridView ID="GridView1" runat="server"></asp:GridView>
    </div>

    </asp:Content>
