<%@ Page Language="VB" Debug="True" %>
<%@ import Namespace="System.Data" %>
<%@ import Namespace="System.Data.Oledb" %>
<%@ import Namespace="System.Data.SqlClient" %>

<script runat="server">

  Sub Page_Load()

          'Call recipe count
          DisplayRecipeCount()

          'Call count unpprove recipes
          UnApproveRecipe()

          DisplayCategoryCount()

          'Call count total comments
          DisplayCommentsCount()

          GetDropdownlistCategory()

          'Call check user function - Check if user has started a session 
          Check_User()

          'Display admin user name
          lblusername.Text = "Welcome Admin:&nbsp;" & session("userid")

       'Display the name of category sorted and search
      If Request("search") <> "" Then
          lblSortedCat.Text = "Category Sorted:&nbsp;" & Request("search")
       End If

       If Not Page.IsPostBack then

           Application.Lock()
           Application("iSortIndex") = 1
           Application.Unlock()
           GetRecipes("id ASC")

       If Request("search") <> "" Then        
          lblSortedCat.Text = "Category Sorted:&nbsp;" & Request("search")
          
       End If

    End If

 End Sub


  Sub DisplayCategoryCount()

        Dim CmdCount As New OleDbCommand("Select Count(CAT_ID) From RECIPE_CAT", New OleDbConnection(strConnection))
        CmdCount.Connection.Open()
        lbCountCat.Text = "Total Category:&nbsp;" & CmdCount.ExecuteScalar()
        CmdCount.Connection.Close()

   End Sub



  Sub GetRecipes(strSortSQL as string)

         Dim strSearchSQL as string
         Dim strSQL as string

        'It check if it is a search or not
         if Request("search") <> "" then
             strSearchSQL = " WHERE ID LIKE '%" & Replace(Request("search"),"'","''") & "%'"
             strSearchSQL += " OR Name LIKE '%" & Replace(Request("search"),"'","''") & "%'"
             strSearchSQL += " OR Author LIKE '%" & Replace(Request("search"),"'","''") & "%'"
             strSearchSQL += " OR Category LIKE '%" & Replace(Request("search"),"'","''") & "%'"
         else
             strSearchSQL = ""
         end if

        'Creates the SQL statement
         strSQL = "SELECT * FROM Recipes" & strSearchSQL & " ORDER BY " & strSortSQL
         objConnection = New OledbConnection(strConnection)
         objCommand = New OledbCommand(strSQL, objConnection)

         Dim RecipeAdapter as New OledbDataAdapter(objCommand)
         Dim dts as New DataSet()
         RecipeAdapter.Fill(dts)

         Recipes_table.DataSource = dts  
         Recipes_table.DataBind()

        
 End Sub

   

  Sub UnApproveRecipe()

Dim CmdCount As New OleDbCommand("Select Count(ID) From Recipes Where LINK_APPROVED = 0", New OleDbConnection(strConnection))
        CmdCount.Connection.Open()
        lblunapproved.Text = "Recipes Waiting For Approval:&nbsp;" & CmdCount.ExecuteScalar() 
        CmdCount.Connection.Close()

  End Sub




  Sub DisplayCommentsCount()

        Dim CmdCount As New OleDbCommand("Select Count(ID) From COMMENTS_RECIPE", New OleDbConnection(strConnection))
        CmdCount.Connection.Open()
        lbCountComments.Text = "Total Comments:&nbsp;" & CmdCount.ExecuteScalar()
        CmdCount.Connection.Close()

  End Sub




   Sub DisplayRecipeCount()

        Dim CmdCount As New OleDbCommand("Select Count(ID) From Recipes", New OleDbConnection(strConnection))
        CmdCount.Connection.Open()
        lbCountRecipe.Text = "Total Recipes:&nbsp;" & CmdCount.ExecuteScalar()
        CmdCount.Connection.Close()

     End Sub



     Sub SearchRecipe(sender as object, e as EventArgs)

         Dim address as string
         address = "recipemanager.aspx?&search=" & Request("search")
         Server.Transfer(address)

     End Sub



     Sub GetCatName(sender as object, e as EventArgs)

         Dim address as string
         address = "recipemanager.aspx?search=" & Request("CategoryName")
         Server.Transfer(address)

     End Sub


     Sub Edit_Handle(sender as Object, e As DataGridCommandEventArgs)

        If (e.CommandName="edit") then
            Dim iIdNumber as TableCell = e.Item.Cells(1)
            Dim address as string

            address = "editing.aspx?id=" & iIdNumber.Text
            Server.Transfer(address)

        End if

  End Sub



 'Display category name in the dropdownlist
 Sub GetDropdownlistCategory()

   Dim myConnection as New OledbConnection(strConnection)

    Dim strSQL as String = "SELECT CAT_ID, CAT_TYPE From RECIPE_CAT"
                             
    Dim myCommand as New OledbCommand(strSQL, myConnection)

	myConnection.Open()
	
	Dim objDR as OledbDataReader
	objDR = myCommand.ExecuteReader(CommandBehavior.CloseConnection)
	
	'Databind the DataReader to the listbox Web control
	CategoryName.DataSource = objDR
	CategoryName.DataBind()
	
	'Add a new listitem to the beginning of the listitemcollection
	CategoryName.Items.Insert(0, new ListItem("Choose Category to Display"))

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
<h3>Recipe Manager</h3>
<asp:Label ID="lblusername" runat="server" />
<br />
<br />
<asp:Label ID="lbCountRecipe" runat="server" />
<br />
<br />
<asp:Label ID="lblunapproved" runat="server" />&nbsp;-&nbsp;
<asp:HyperLink tooltip="Click this link to approve recipe" runat="server" ID="approvallink" NavigateUrl="recipeapproval.aspx">Recipe Approval Manager Page</asp:HyperLink> 
<br />
<br />
<asp:Label ID="lbCountComments" runat="server" />&nbsp;-&nbsp;
<asp:HyperLink tooltip="Click this link to edit/delete recipe comments" runat="server" ID="countcommentlink" NavigateUrl="commentsmanager.aspx">Recipe Comments Manager Page</asp:HyperLink>
<br />
<br />
<asp:Label ID="lbCountCat" runat="server" />&nbsp;-&nbsp;
<asp:HyperLink tooltip="Click this link to edit/delete and add a recipe category" runat="server" ID="editcat" NavigateUrl="categorymanager.aspx">Recipe Category Manager Page</asp:HyperLink>  
<br />
<br />
</div>
<form style="margin-top: 3px; margin-bottom: 1px;" runat="server">
<table border="0" cellpadding="2" width="90%" align="center" bgcolor="#F7F7F7">
  <tr>
    <td width="20%" bgcolor="#FDFDFD" class="content2">Search recipe by ID, Name or Author:</td>
    <td width="80%" bgcolor="#FDFDFD"><asp:TextBox runat="server" class="textbox" id="search" maxlength="25" size="25"/> 
    <br />
    <asp:Button runat="server" OnSubmit="SearchRecipe" class="submit" Text="Search"/>
</td>
  </tr>
</table>
<table width="100%" border="0" cellspacing="1">
  <tr>
    <th scope="row"><div style="text-align: left; padding-left: 25px; margin-top: 12px;">
<span class="content2"><b>Sort Category:</b></span><asp:listbox id="CategoryName" runat="server" Rows="1" 
               DataTextField="CAT_TYPE" DataValueField="CAT_TYPE" />
<asp:Button runat="server" ID="GO" OnClick="GetCatName" class="submit" Text="Go"/>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<asp:HyperLink runat="server" NavigateUrl="recipemanager.aspx" class="content2">Back to Default View</asp:HyperLink>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<asp:Label font-name="verdana" font-size="9" ID="lblSortedCat" runat="server" />
</div>
</th>
  </tr>
  <tr>
    <th scope="row"><div align="left">
     <asp:DataGrid runat="server" id="Recipes_table" cssclass="hlink" AutoGenerateColumns="False" AllowSorting="true"
     Backcolor="#f7f7f7" BorderStyle="none" BorderColor="#ffffff" cellpadding="5" Width="95%" HorizontalAlign="Center" PageSize="30" onSortCommand="Sort_Recipes" AllowPaging="True" OnPageIndexChanged="New_Page" onItemCommand="Edit_Handle"> 
     <HeaderStyle Font-Bold="True" BackColor="#6898d0" cssclass="header" />
     <AlternatingItemStyle BackColor="White" />                                   
     <Columns>
     <asp:ButtonColumn Text="Edit..." HeaderText="Edit" CommandName="edit" />
     <asp:BoundColumn DataField="ID" HeaderText="ID" SortExpression="id ASC" />   
     <asp:HyperLinkColumn DataTextField="Name" HeaderText="Name" SortExpression="Name" DataNavigateUrlField="ID" DataNavigateUrlFormatString="javascript:var w=window.open('viewing.aspx?id={0}','','width=700,height=690,scrollbars=yes');" />
     <asp:BoundColumn DataField="Category" HeaderText="Category" SortExpression="Category" />
     <asp:BoundColumn DataField="Author" HeaderText="Author" SortExpression="Author" />
     <asp:BoundColumn DataField="Date" HeaderText="Date" SortExpression="Date" />
     <asp:BoundColumn DataField="Hits" HeaderText="Hits" SortExpression="Hits" />
     <asp:BoundColumn DataField="Rating" HeaderText="Rating" SortExpression="Rating" /> 
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