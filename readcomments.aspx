<%@ Page Language="VB" Debug="true" %>
<%@ import Namespace="System.Data" %>
<%@ import Namespace="System.Data.Oledb" %>

<script runat="server">

 'Handle page load event
 Sub Page_Load(Sender As Object, E As EventArgs)

          
           'Call the bindlist - show data
           BindList()


 End Sub



 'Bind data
 Sub BindList()
        
 strSQL = "SELECT * From COMMENTS_RECIPE Where ID=" & Request.QueryString("id") & " Order By Date Desc"

         'Call Open database - connect to the database
         DBconnect()

         Dim RecipeAdapter as New OledbDataAdapter(objCommand)
         Dim dts as New DataSet()
         RecipeAdapter.Fill(dts, "ID")

         RecComments.DataSource = dts         
         RecComments.DataBind() 

  End Sub


 'Database connection string - Open database
 Sub DBconnect()

     objConnection = New OledbConnection(strConnection)
     objCommand = New OledbCommand(strSQL, objConnection)

 End Sub

    'Declare public so it will accessible in all subs
    Public strDBLocation = DB_Path()
    Public strConnection as string = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strDBLocation
    Public objConnection
    Public objCommand
    Public strSQL as string


</script>

<!--#include file="inc_databasepath.aspx"-->      

<!--#include file="inc_header.aspx"-->

<table border="0" cellpadding="0" cellspacing="0" align="center" width="70%">
  <tr>
    <td width="100%"></td>
  </tr>
  <tr>
    <td width="100%">
<asp:DataList width="100%" id="RecComments" RepeatColumns="1" runat="server">
      <ItemTemplate>
    <div class="divwrap">
<div class="divbd">
Author:&nbsp;<%# DataBinder.Eval(Container.DataItem, "AUTHOR") %>
<br />
Email:&nbsp;<%# DataBinder.Eval(Container.DataItem, "EMAIL") %>
<br />
Date:&nbsp; <%# FormatDateTime(DataBinder.Eval(Container.DataItem, "DATE"),vbShortDate) %>
<br />
Comment:
<br /> 
<%# Replace(DataBinder.Eval(Container.DataItem, "COMMENTS"), Chr(13), "<br>") %>
     </div>
   </div>
 </ItemTemplate>
</asp:DataList>
</td>
  </tr>
</table>

<!--#include file="inc_footer.aspx"-->