<%@ Page Language="VB" Debug="true" %>
<%@ import Namespace="System.Data" %>
<%@ import Namespace="System.Data.Oledb" %>

<script runat="server">

 'Handle page load event
 Sub Page_Load(Sender As Object, E As EventArgs)
              

 strSQL = "Update Recipes set LINK_APPROVED = 1 where id=" & Request.QueryString("id")

            'Call Open database - connect to the database
            DBconnect()

            objConnection.Open()
            objCommand.ExecuteNonQuery()
            
            'Close the db connection and free up memory
            DBclose()

            Dim address as string

            address = "recipeapproval.aspx"
            Server.Transfer(address)

  End Sub


 'Database connection string - Open database
 Sub DBconnect()

     objConnection = New OledbConnection(strConnection)
     objCommand = New OledbCommand(strSQL, objConnection)

 End Sub


 'Close the db connection and free up memory
 Sub DBclose()

    objCommand = nothing
    objConnection.Close()
    objConnection = nothing

 End Sub

    'Declare public so it will accessible in all subs
    Public strDBLocation = DB_Path()
    Public strConnection as string = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strDBLocation
    Public objConnection
    Public objCommand
    Public strSQL as string

</script>

<!--#include file="inc_databasepath.aspx"-->      

<html>

<head><title>Printing Recipe</title>
<style type="text/css" media="screen">@import "cssreciaspx.css";</style>
</head>

<body>
<div style="margin: 50px;">
<table border="0" cellpadding="5" align="center" cellspacing="5" width="70%">
<tr>
    <td width="100%">
</td>
  </tr>
</table>
</div>

</body>

</html>