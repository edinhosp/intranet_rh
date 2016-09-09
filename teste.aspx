<%@ Page language="c#" Codebehind="HtmlControlsDemo.aspx.cs"
AutoEventWireup="false" Inherits="Html_Web_Controls.HtmlControlsDemo"
Trace="True"%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN" >
<HTML runat="server" id="htmltag">
  <HEAD id="headtag" runat="server">
    <title id="titletag" runat="server">HtmlControls</title>
</HEAD>
  <body >
    <h1 runat="server" id="h1Title">HtmlControls</h1>
    <form id="Form1" method="post" runat="server">
        <table runat="server" id="table1">
        <tr runat="server" id="row1">
            <td runat="server" id="cell1">
                <span runat="server" id="span1">Span</span>
            </td>
            <td runat="server" id="cell2">
                    <input id="input1" runat="server"
                        value="input1">
            </td>
            <td runat="server" id="cell3">
                <input id="radio1" runat="server"
                    type="radio" value="Choice 1" >
                <input id="radio2" runat="server"
                    type="radio" value="Choice 2" >
            </td>
            <td runat="server" id="cell4"><input type="button"
            runat="server" id="button1" value="Button" ></td>
        </tr>
        <tr runat="server" id="row2">
            <td runat="server" id="cell5">
                <select runat="server" id="select1">
                    <option >Option 1</option>
                    <option >Option 2</option>
                </select>
            </td>
            <td runat="server" id="cell6">
                <textarea runat="server" id="textarea1"></textarea>
            </td>
            <td runat="server" id="cell7">
                <a runat="server" id="anchor1"
                    href="http://aspsmith.com/">
                    <img
                    src="http://aspsmith.com/images/aspsmith_logo_197x53.gif"
                        runat="server" id="image1" >
                </a>
            </td>
            <td runat="server" id="cell8">
                <input type="checkbox" runat="server"
                    id="check1" value="choice 1" name="check1">
                <input type="checkbox" runat="server"
                    id="check2" value="choice 2" name="check2">
            </td>
        </tr>
        </table>
        <input type="hidden" runat="server" id="hidden1">
        <input type="file" runat="server" id="file1" >
        <input type="image" runat="server" id="imagebutton1"
        alt="Input type=image"
        src="http://aspsmith.com/images/aspsmith_logo_197x53.gif">
    </form>

  </body>
</HTML>

