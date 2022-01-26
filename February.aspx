<%@ Page Title="" Language="C#" MasterPageFile="~/Site.Master" AutoEventWireup="true" CodeBehind="February.aspx.cs" Inherits="_Git_.February" %>

<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <div id="content-wrapper" class="d-flex flex-column">
        <!-- Main Content -->
        <div id="content" class="mt-5 ml-5 mr-5">
            <h3>
                <center>
                    February
                </center>
            </h3>
            <hr />
            <div class="ml-5 mt-5">
                <h4>Import Excel</h4>
                <h6>Add new information by uploading file below.</h6>
                <h6>The file format should be in .xls format. </h6>
                <br />
                <div class="card-excelupload">
                    <p>Please drag/upload file here.</p>
                    <div class="excel-fileupload">
                        <asp:FileUpload class="btn btn-light shadow m-3" ID="FileUploadExcel" runat="server" />
                        <asp:Button class="btn btn-primary shadow" ID="btnUpload" runat="server" OnClick="UploadExcel_Click" Text="Upload File" />
                        <asp:Label ID="UploadLabel" runat="server" Text=" "></asp:Label>
                    </div>
                </div>
            </div>
        </div>
</asp:Content>
