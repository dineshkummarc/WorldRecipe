<%@ Page Language="VB" Debug="True"%>
<%@ Import Namespace="System.Web.Mail" %>

<script language="VB" runat="server">

  Sub Page_Load(Source As Object, E As EventArgs)

    If Not Page.IsPostBack Then
      
      Dim strName as string = Request.QueryString("n")
      Dim strCat as string = Request.QueryString("c")
      Dim strBody As String

      strBody = "Hi Friend," & vbCrLf & vbCrLf _
            & "I thought you might be interested in this recipe I found at www.ex-designz.net:" & vbCrLf & vbCrLf _
            & "Recipe Name: " & strName & vbCrLf _
            & "Category: " & strCat & vbCrLf _
            & vbCrLf _
        & Request.QueryString("url") & vbCrLf

      txtMessage.Text = strBody
    End If
  End Sub


  Sub btnSendMsg_OnClick(Source As Object, E As EventArgs)

    Dim Messagesnd As New MailMessage
    Dim sndingmail As SmtpMail

    If Page.IsValid() Then

      Messagesnd.From    = txtFromEmail.Text
      Messagesnd.To      = txtToEmail.Text
      Messagesnd.Bcc     = "extremedexter_z2001@yahoo.com"
      Messagesnd.Subject = txtFromName.Text & " has emailed you " & Request.QueryString("n") & " recipe"
      Messagesnd.Body    = txtMessage.Text & vbCrLf _
        & "This message was sent from: " _
        & Request.ServerVariables("SERVER_NAME") & "." _
        & vbCrLf & vbCrLf _
        & "This email was sent to " & txtToEmail.Text & "."

      ' SMTP server's name,localhost or ip address!
      sndingmail.SmtpServer = "localhost"
      sndingmail.Send(Messagesnd)

      Panel1.Visible = False

      lblsentmsg.Text = "Your message has been sent to " _
	    & txtToEmail.Text & "."
    End If

  End Sub

</script>


<html>
<head>
<title>Sending Recipe To a Friend</title>
<style type="text/css" media="screen">@import "cssreciaspx.css";</style>
</head>
<body>
<asp:Panel ID="Panel1" runat="server">
<form runat="server">
<table align="center" cellspacing="0" cellpadding="0" width="40%">
<tr><td>
<br />
<div align="center"><h2>Sending <%=Request.QueryString("n")%> Recipe to a Friend</h2></div>
<table border="0" cellspacing="1" cellpadding="1" width="100%">
  <tr>
    <td valign="top" align="right" class="content6"><b>Your Name:</b></td>
    <td>
      <asp:TextBox id="txtFromName" size="25" cssClass="textbox" runat="server" />
      <asp:RequiredFieldValidator runat="server"
        id="validNameRequired" ControlToValidate="txtFromName"
        cssClass="cred2" errormessage="* Name:<br />"
        display="Dynamic" />
    </td>
  </tr>
  <tr>
    <td valign="top" align="right" class="content6"><b>Your Email:</b></td>
    <td>
      <asp:TextBox id="txtFromEmail" size="25" cssClass="textbox" runat="server" />
      <asp:RequiredFieldValidator runat="server"
        id="validFromEmailRequired" ControlToValidate="txtFromEmail"
        cssClass="cred2" errormessage="* Email:<br />"
        display="Dynamic" />
      <asp:RegularExpressionValidator runat="server"
        id="validFromEmailRegExp" ControlToValidate="txtFromEmail"
        ValidationExpression="^[\w-]+@[\w-]+\.(com|net|org|edu|mil)$"
        cssClass="cred2" errormessage="Not Valid"
        Display="Dynamic" />    
    </td>
  </tr>
<tr>
    <td valign="top" align="right" class="content6"><b>Friend's Name:</b></td>
    <td>
      <asp:TextBox id="toname" size="25" cssClass="textbox" runat="server" />
      <asp:RequiredFieldValidator runat="server"
        id="validFriendNameRequired" ControlToValidate="toname"
        cssClass="cred2" errormessage="* Name:<br />"
        display="Dynamic" />
    </td>
  </tr>
  <tr>
    <td valign="top" align="right" class="content6"><b>Friend's Email:</b></td>
    <td>
      <asp:TextBox id="txtToEmail" size="25" cssClass="textbox" runat="server" />
      <asp:RequiredFieldValidator runat="server"
        id="validToEmailRequired" ControlToValidate="txtToEmail"
        cssClass="cred2" errormessage="* Email:<br />"
        display="Dynamic" />
      <asp:RegularExpressionValidator runat="server"
        id="validToEmailRegExp" ControlToValidate="txtToEmail"
        ValidationExpression="^[\w-]+@[\w-]+\.(com|net|org|edu|mil)$"
        cssClass="cred2" errormessage="Not Valid:"
        Display="Dynamic" />  
    </td>
  </tr>
  <tr>
    <td colspan="2">
      <asp:TextBox id="txtMessage" cssClass="textbox" Cols="50" TextMode="MultiLine"
        Rows="10" ReadOnly="True" runat="server" />
      <br />
    <asp:Button id="btnSend" cssClass="submit" Text="Send Recipe"
      OnClick="btnSendMsg_OnClick" runat="server" />
    </td>
  </tr>
</table>
</td></tr>
</table>
</form>
</asp:Panel>
<div style="text-align: center;" class="content2"><asp:HyperLink runat="server" NavigateUrl="JavaScript:onClick= window.close()" class="content2">Close Window</asp:HyperLink></div>
<br />
<br />
<asp:Label cssClass="content2" id="lblsentmsg" runat="server" />
</body>
</html>