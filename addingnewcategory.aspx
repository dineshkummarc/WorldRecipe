<%@ Page Language="VB" Debug="true" %>
<%@ import Namespace="System.Data" %>
<%@ import Namespace="System.Data.Oledb" %>

<script runat="server">

   Sub Page_Load()
    
    
            Check_User()
    
 End Sub
    
   

   'It add a new category 
    Sub Add_Category(sender As Object, e As System.EventArgs)
    
        Dim strSQL as string
    
        objConnection = New OledbConnection(strConnection)
        objConnection.Open()
    
    
        strSQL = "insert into RECIPE_CAT (CAT_TYPE) values ('" & replace(request("Category"),"'","''") & "')"
    
        objCommand = New OledbCommand(strSQL,objConnection)
        objCommand.ExecuteNonQuery()
    
        objCommand = nothing
        objConnection.Close()
        objConnection = nothing
        response.redirect("categorymanager.aspx")
    
    End Sub


</script>

<!--#include file="config.aspx"-->

<!--Powered By www.Ex-designz.net Recipe Cookbook ASP.NET version - Author: Dexter Zafra, Norwalk,CA-->

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title>Edit - Delete Page - www.ex-designz.net</title>
<style type="text/css" media="screen">@import "cssreciaspx.css";</style>
</head>
<body>
<br />
<br />
<br />
<div align="center"><span class="hlink">To change the category name, type in the new category name in the field and click the update button. 
<br />
To delete this category, click the delete button and you will be redirected to the delete category page.
</span></div>
<form runat="server">

                       <table width="40%" border="0" cellpadding="3" cellspacing="0" align="center">
            		                  <tr>
            			                 <td colspan=2  bgcolor="#6898d0">
            <span class="content3">Adding New Category</span>
            			                 </td>
            		                  </tr>
            		                  <tr>
          <td bgcolor="#f7f7f7" class="content2">Enter Category Name:</td>   		
         <td bgcolor="#fbfbfb">
        <asp:TextBox runat="server" id="Category" class="textbox" size="25" maxlenght="25" />
            			                 </td>
            		                  </tr>                 		                 
            		                  <tr>
            			       <td align=center colspan=2 bgcolor="#ffffff">
       <asp:Button runat="server" Text="Update" id="Addbutton" class="submit" onclick="Add_Category" />
            			             </td>
            		                  </tr>
            	                   </table>
        </form>
<br />
<div align="center"><asp:HyperLink cssClass="dt" tooltip="Back to Recipe Manager Main Page" runat="server" ID="approvallink" NavigateUrl="recipemanager.aspx">Recipe Manager Main Page</asp:HyperLink></div>
    </body>
</html>
