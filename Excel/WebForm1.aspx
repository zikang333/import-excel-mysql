<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="WebForm1.aspx.cs" Inherits="Excel.WebForm1" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
    <style type="text/css">
        .auto-style2 {
            margin-left: 0px;
        }

        .auto-style5 {
            height: 715px;
        }

        .auto-style6 {
            width: 100%;
            height: 106px;
        }

        .auto-style10 {
            width: 141px;
            height: 37px;
        }

        .auto-style11 {
            height: 37px;
        }

        .auto-style12 {
            width: 141px;
            height: 42px;
        }

        .auto-style13 {
            height: 42px;
        }

        .auto-style14 {
            width: 141px;
            height: 45px;
        }

        .auto-style15 {
            height: 45px;
        }

        .auto-style16 {
            width: 141px;
            height: 59px;
        }

        .auto-style17 {
            height: 59px;
        }

        .Grid {
            background-color: #fff;
            margin: 5px 0 10px 0;
            border: solid 1px #525252;
            border-collapse: collapse;
            font-family: Calibri;
            color: #474747;
        }

            .Grid td {
                padding: 2px;
                border: solid 1px #c1c1c1;
            }

            .Grid th {
                padding: 4px 2px;
                color: #fff;
                background: #363670 url(Images/grid-header.png) repeat-x top;
                border-left: solid 1px #525252;
                font-size: 0.9em;
            }

            .Grid .alt {
                background: #fcfcfc url(Images/grid-alt.png) repeat-x top;
            }

            .Grid .pgr {
                background: #363670 url(Images/grid-pgr.png) repeat-x top;
            }

                .Grid .pgr table {
                    margin: 3px 0;
                }

                .Grid .pgr td {
                    border-width: 0;
                    padding: 0 6px;
                    border-left: solid 1px #666;
                    font-weight: bold;
                    color: #fff;
                    line-height: 12px;
                }

                .Grid .pgr a {
                    color: Gray;
                    text-decoration: none;
                }

                    .Grid .pgr a:hover {
                        color: #000;
                        text-decoration: none;
                    }

        .button {
            transition-duration: 0.4s;
        }

            .button:hover {
                background-color: #4CAF50; /* Green */
                color: white;
            }

        .button2 {
            transition-duration: 0.4s;
        }

            .button2:hover {
                background-color: #3e80f7; /* Blue */
                color: white;
            }

        .auto-style18 {
            margin-left: 16px;
        }

    </style>

</head>
<body style="height: 749px">
    <form id="form1" runat="server" class="auto-style5">
        <table class="auto-style6">
            <tr>
                <td class="auto-style10">
                    <asp:Label ID="Label1" runat="server" Text="Choose Your File"></asp:Label>
                </td>
                <td class="auto-style11">
                    <asp:Button ID="Button2" runat="server" CssClass="button2" OnClick="Browse_Click" Text="Browse" Width="85px" Height="30px"/>
                </td>
                <td class="auto-style11"></td>
            </tr>
            <tr>
                <td class="auto-style12">
                    <asp:Label ID="Label2" runat="server" Text="File Path"></asp:Label>
                </td>
                <td class="auto-style13">
                    <asp:TextBox ID="TextBox1" runat="server" CssClass="auto-style2" Width="517px" Height="21px"></asp:TextBox>
                    <asp:Button ID="Button4" runat="server" CssClass="button2" Height="30px" OnClick="Upload_Click" Text="Upload" Width="85px" />
                </td>
                <td class="auto-style13"></td>
            </tr>
            <tr>
                <td class="auto-style16">
                    <asp:Label ID="Label4" runat="server" Text="Title"></asp:Label>
                </td>
                <td class="auto-style17">
                    <asp:DropDownList ID="DropDownList1" runat="server" Height="40px" OnSelectedIndexChanged="DropDownList1_SelectedIndexChanged" Width="525px" AutoPostBack="True">
                    </asp:DropDownList>
                </td>
                <td class="auto-style17"></td>
            </tr>
            <tr>
                <td class="auto-style16">
                    <asp:Label ID="Label5" runat="server" Text="ISKU"></asp:Label>
                </td>
                <td class="auto-style17">
                    <asp:TextBox ID="TextBox2" runat="server" Width="193px" BorderStyle="Groove" Enabled="False"></asp:TextBox>
                </td>
                <td class="auto-style17">&nbsp;</td>
            </tr>
            <tr>
                <td class="auto-style14">Overview</td>
                <td class="auto-style15">
                    <asp:Button ID="Button3" runat="server" Height="30px" OnClick="Check_Click" Text="Check" Width="85px" CssClass="button2" />
                </td>
                <td class="auto-style15">&nbsp;</td>
            </tr>
            <tr>
                <td class="auto-style14">
                    <asp:Label ID="Label3" runat="server" Text="Import to DB"></asp:Label>
                </td>
                <td class="auto-style15">
                    <asp:Button ID="Button1" runat="server" OnClick="Import_Click" Text="Import" CssClass="button" Height="30px" Width="85px" />
                </td>
                <td class="auto-style15">&nbsp;</td>
            </tr>
        </table>
        <asp:GridView ID="GridView1" runat="server" Width="255px" CssClass="Grid" AlternatingRowStyle-CssClass="alt" PagerStyle-CssClass="pgr" ItemStyle-HorizontalAlign="Right">
            <RowStyle HorizontalAlign="Center" VerticalAlign="Middle" />
        </asp:GridView>
        <br />
        <br />
        <br />
        <br />
        <br />
        <br />
        <br />
        <br />
        <br />
    </form>
</body>
</html>
