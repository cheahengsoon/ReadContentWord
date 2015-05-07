<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="Default.aspx.cs" Inherits="RWFile.Default" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
  <title>Untitled Page</title>
</head>
<body>
  <form id="form1" runat="server">
      <div>
        <div>
<h2>
Read Word Files in ASP.Net,C#.Net & VB.Net</h2>
<table>
    <tbody>
<tr>
        <td>Select Word File
        </td>
        <td><asp:fileupload borderstyle="None" id="fuWordFile" runat="server" width="549px"></asp:fileupload></td>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
    </tr>
<tr>
        <td></td>
        <td><asp:button id="btnUpload" onclick="btnUpload_Click" runat="server" text="Upload"></asp:button></td>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
    </tr>
<tr>
        <td colspan="2">Content of the word file is given below:-
        </td>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
        <td>
       <br />
             </td>
    </tr>
<tr>
        <td colspan="5">
          
            <asp:TextBox ID="txtintro" runat="server" Height="141px" Width="357px" TextMode="MultiLine"></asp:TextBox>
            <asp:TextBox ID="txtbenefit" runat="server" Height="141px" Width="350px" style="margin-top: 0px" TextMode="MultiLine"></asp:TextBox>
             <asp:TextBox ID="txtcontenttitle" runat="server" Height="143px" Width="348px" TextMode="MultiLine"></asp:TextBox>
            
        </td>
    </tr>
<tr>
        <td colspan="2">&nbsp;</td>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
        <td>
            &nbsp;</td>
    </tr>
<tr>
        <td colspan="2" rowspan="3"><asp:textbox height="500px" id="TextBox1" runat="server" textmode="MultiLine" width="100%"></asp:textbox>
        </td>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
        <td rowspan="3">
          
            <br />
            <br />
            
        </td>
    </tr>
<tr>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
    </tr>
<tr>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
    </tr>
</tbody></table>
</div>
      </div>
  </form>
</body>
</html>