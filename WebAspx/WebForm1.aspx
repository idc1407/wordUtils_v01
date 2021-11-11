<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="WebForm1.aspx.cs" Inherits="WebAspx.WebForm1" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
    <script src="Scripts/jquery-3.4.1.min.js"></script>
    <link href="Content/bootstrap.min.css" rel="stylesheet" />
    <script src="customScripts/WordUtil.js"></script>
</head>
<body>
    <div class="container body-content">
        <div class="row">
            <form id="form1" runat="server" class="form-horizontal">
                <div class="row">
                    <div class="form-group">
                        <div class="col-md-offset-2 col-md-10">
                            <asp:FileUpload ID="FileUpload1" runat="server" Class="form-control" required />
                        </div>
                    </div>
                </div>
                <div class="row">
                    <div class="form-group">
                        <div class="col-md-2">
                        </div>
                        <div class="col-md-2">
                            <asp:CheckBox ID="IsFooterTextChange" Text="&nbsp;&nbsp;Footer Text Change" runat="server" />
                        </div>
                        <div class="col-md-3">
                            <label for="FooterTextFind" class="control-label">Find Text </label>
                            <asp:TextBox ID="FooterTextFind" runat="server" class="form-control"
                                placeholder="Find Text" />
                        </div>
                        <div class="col-md-3">
                            <label for="FooterTextReplace" class="control-label">Replace Text</label>
                            <asp:TextBox ID="FooterTextReplace" runat="server" class="form-control"
                                placeholder="Replace Text" />
                        </div>
                        <div class="col-md-2"></div>
                    </div>
                </div>

                <div class="row">
                    <div class="form-group">
                        <div class="col-md-2">
                        </div>
                        <div class="col-md-2">
                            <asp:CheckBox ID="IsHeaderTextChange" Text="&nbsp;&nbsp;Header Text Change" runat="server" />
                        </div>
                        <div class="col-md-3">
                            <label for="HeaderTextFind" class="control-label">Find Text </label>
                            <asp:TextBox ID="HeaderTextFind" runat="server" class="form-control"
                                placeholder="Find Text" />
                        </div>
                        <div class="col-md-3">
                            <label for="HeaderTextReplace" class="control-label">Replace Text</label>
                            <asp:TextBox ID="HeaderTextReplace" runat="server" class="form-control"
                                placeholder="Replace Text" />
                        </div>
                        <div class="col-md-2"></div>
                    </div>
                </div>

                <div class="row">
                    <div class="form-group">
                        <div class="col-md-2"></div>
                        <div class="col-md-3">
                            <asp:CheckBox ID="IsBalanceSheetTableDelete" Text="&nbsp;&nbsp;Balance Sheet Table Delete" runat="server" />
                        </div>
                        <div class="col-md-3">
                            <asp:CheckBox ID="IsOtherOptionA" Text="&nbsp;&nbsp;Other Options A" runat="server" />
                        </div>
                        <div class="col-md-3">
                            <asp:CheckBox ID="IsOtherOptionB" Text="&nbsp;&nbsp;Other Options B" runat="server" />
                        </div>
                        <div class="col-md-1"></div>
                    </div>
                </div>

                <div class="row">
                    <div class="form-group">
                        <div class="text-center col-md-12">
                            <asp:Button ID="process" runat="server" Text="Proceed" OnClick="process_Click" Class="btn btn-primary" />
                        </div>
                    </div>
                </div>

                <div class="row">
                    <div class="form-group">
                        <div class="text-center col-md-12">
                            <asp:Button ID="Button1" Text="Ajax" Class="btn btn-primary" />
                        </div>
                    </div>
                </div>



            </form>
        </div>

        <div class="row text-center col-md-12" id="smessage">
            <h2>
                <asp:Label ID="smessage" class="label label-success" runat="server"></asp:Label></h2>
        </div>


    </div>
</body>
</html>
