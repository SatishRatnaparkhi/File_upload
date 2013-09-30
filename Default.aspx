<%@ Page Language="C#" AutoEventWireup="true" CodeFile="Default.aspx.cs" Inherits="_Default" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
</head>
<body>
    <form id="form1" runat="server">
        <div>
            <asp:FileUpload ID="fuCSV" runat="server"></asp:FileUpload>

            <asp:Button ID="btnUpload" runat="server" OnClick="btnUpload_Click" Text="Import/Insert" />
            <asp:Label runat="server" ID="lblerror"></asp:Label>
        </div>
    </form>
</body>
</html>
