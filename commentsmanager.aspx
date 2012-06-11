<%@ Page Language="VB" Debug="True" %>
<%@ import Namespace="System.Data" %>
<%@ import Namespace="System.Data.Oledb" %>
<%@ import Namespace="System.Data.SqlClient" %>

<script runat="server">

  Sub Page_Load()

          'Call recipe count
          DisplayCommentsCount()

          'Call check user function - Check if user has started a session 
          Check_User()

          'Display admin user name
          lblusername.Text = "Welcome Admin:&nbsp;" & session("userid")

       If Not Page.IsPostBack then

           Application.Lock()
           Application("iSortIndex") = 1
           Application.Unlock()
           GetRecipes("DATE DESC")

    End If

 End Sub



  Sub GetRecipes(strSortSQL as string)

         Dim strSQL as string

        'Creates the SQL statement
         strSQL = "SELECT ID,AUTHOR,EMAIL,DATE FROM COMMENTS_RECIPE Order By DATE DESC"
         objConnection = New OledbConnection(strConnection)
         objCommand = New OledbCommand(strSQL, objConnection)

         Dim RecipeAdapter as New OledbDataAdapter(objCommand)
         Dim dts as New DataSet()
         RecipeAdapter.Fill(dts)

         Recipes_table.DataSource = dts  
         Recipes_table.DataBind()

        
 End Sub



  Sub DisplayCommentsCount()

        Dim CmdCount As New OleDbCommand("Select Count(ID) From COMMENTS_RECIPE", New OleDbConnection(strConnection))
        CmdCount.Connection.Open()
        lbCountRecipe.Text = "Total Comments:&nbsp;" & CmdCount.ExecuteScalar()
        CmdCount.Connection.Close()

   End Sub




  Sub Edit_Handle(sender as Object, e As DataGridCommandEventArgs)

        If (e.CommandName="edit") then
            Dim iIdNumber as TableCell = e.Item.Cells(1)
            Dim address as string

            address = "editingcomments.aspx?id=" & iIdNumber.Text
            Server.Transfer(address)

        End if

  End Sub




  Sub Sort_Recipes(sender As Object, e As DataGridSortCommandEventArgs)

         Dim SortExprs() As String
         Dim CurrentSearchMode As String, NewSearchMode As String
         Dim ColumnToSort As String, NewSortExpr as String

         SortExprs = Split(e.SortExpression, " ")
         ColumnToSort = SortExprs(0)

         If SortExprs.Length() > 1 Then
           CurrentSearchMode = SortExprs(1).ToUpper()
           If CurrentSearchMode = "ASC" Then
              NewSearchMode = "Desc"
           Else
              NewSearchMode = "Asc"
           End If
         Else
           NewSearchMode = "Desc"
         End If

         NewSortExpr = ColumnToSort & " " & NewSearchMode

         Dim iIndex As Integer

         Select case ColumnToSort.toUpper()
              case "ID"
                 iIndex = 1
              case "Name"
                 iIndex = 2
              case "Category"
                 iIndex = 3
              case "Author"
                 iIndex = 4
              case "Hits"
                 iIndex = 5
              case "Date"
                 iIndex = 6
              case "Rating"
                 iIndex = 7
         End Select

         Application.Lock()
         Application("iSortIndex") = iIndex
         Application.Unlock()

         Recipes_table.Columns(iIndex).SortExpression = NewSortExpr

         GetRecipes(NewSortExpr)

  End Sub



  'Handles page change links - paging system
  Sub New_Page(sender As Object, e As DataGridPageChangedEventArgs)

         Dim iSort
         Application.Lock()
         iSort = Application("iSortIndex")
         Application.Unlock()
         Dim strSortVars = Recipes_table.Columns(iSort).SortExpression
         Recipes_table.CurrentPageIndex = e.NewPageIndex
         GetRecipes(strSortVars)

  End Sub


</script>

<!--#include file="config.aspx"-->

<!--Powered By www.Ex-designz.net Recipe Cookbook ASP.NET version - Author: Dexter Zafra, Norwalk,CA-->

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title>Recipe Manager - www.ex-designz.net</title>
<style type="text/css" media="screen">@import "cssreciaspx.css";</style>
</head>
<body>
<div class="div2">
<h3>Recipe Comments Manager</h3>
<asp:Label ID="lblusername" runat="server" />
<br />
<br />
How to?: 
<br />
To review a comment, click the Date link.
<br />
To view the recipe associated with the comment, click the Author Name link.
<br />
To edit/delete a comment, click the ID link.
<br />
<br />
<asp:Label ID="lbCountRecipe" runat="server" />&nbsp;&nbsp;&nbsp;<asp:Label ID="lblSortedCat" runat="server" />
<br />
<br />
<asp:HyperLink tooltip="Back to Recipe Manager Main Page" runat="server" ID="approvallink" NavigateUrl="recipemanager.aspx">Recipe Manager Main Page</asp:HyperLink>
</div>
<form style="margin-top: 3px; margin-bottom: 1px;" runat="server" action="recipemanager.aspx">
<table width="100%" border="0" cellspacing="1">
  <tr>
    <th scope="row"><div style="text-align: left; padding-left: 25px; margin-top: 12px;"></div></th>
  </tr>
  <tr>
    <th scope="row"><div align="left">
     <asp:DataGrid runat="server" id="Recipes_table" cssclass="hlink" AutoGenerateColumns="False" AllowSorting="true"
     Backcolor="#f7f7f7" BorderStyle="none" BorderColor="#ffffff" cellpadding="5" Width="95%" HorizontalAlign="Center" PageSize="30" onSortCommand="Sort_Recipes" AllowPaging="True" OnPageIndexChanged="New_Page" onItemCommand="Edit_Handle"> 
     <HeaderStyle Font-Bold="True" BackColor="#6898d0" cssclass="header" />
     <AlternatingItemStyle BackColor="White" />                                   
     <Columns>
<asp:HyperLinkColumn DataTextField="ID" HeaderText="ID" SortExpression="ID" DataNavigateUrlField="ID" DataNavigateUrlFormatString="javascript:var w=window.open('editingcomments.aspx?id={0}','','width=650,height=680,resizable=yes,scrollbars=yes');" />
     <asp:HyperLinkColumn DataTextField="AUTHOR" HeaderText="Author" SortExpression="AUTHOR" DataNavigateUrlField="ID" DataNavigateUrlFormatString="javascript:var w=window.open('recipedetail.aspx?id={0}','','width=850,height=700,resizable=yes,scrollbars=yes');" />
<asp:HyperLinkColumn DataTextField="DATE" HeaderText="Date" SortExpression="DATE" DataNavigateUrlField="ID" DataNavigateUrlFormatString="javascript:var w=window.open('readcomments.aspx?id={0}','','width=650,height=700,resizable=yes,scrollbars=yes');" />
     </Columns>
     <PagerStyle Mode="NumericPages" BackColor="#fcfcfc" HorizontalAlign="left" />
    </asp:DataGrid>                                                                                             
   </div></th>
 </tr>
</table>
</form>
<div style="text-align: center; margin-top: 15px;">
<a href="http://www.ex-designz.net" class="hlink" title="Visit our website">Powered By Ex-designz.net World Recipe</a>
</div>
</body>
</html>