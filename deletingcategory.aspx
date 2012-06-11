<%@ Page Language="VB" Debug="true" %>
<%@ import Namespace="System.Data" %>
<%@ import Namespace="System.Data.Oledb" %>

<script runat="server">

   Sub Page_Load()
    

            Dim strSQL as string
    
            Check_User()

            DisplayCount()
    
    	    'Check which action were selected, edit a recipe or delete a recipe
            strSQL = "SELECT CAT_ID, CAT_TYPE From RECIPE_CAT Where CAT_ID =" & request.querystring("catid") 
    
            DataBase_Connect(strSQL)   
            objDataReader.Read()

            'This will be the value to be populated into the textboxes
            CategoryName.text = objDataReader("CAT_TYPE")
    
            DataBase_Disconnect()
    
 End Sub
    


   'Delete the selected comments
   Sub Delete_Category(sender As Object, e As System.EventArgs)
    
        Dim strSQL as string
    
        objConnection = New OledbConnection(strConnection)
        objConnection.Open()
    
        strSQL = "delete * from RECIPE_CAT where CAT_ID = " & request("catid")
        objCommand = New OledbCommand(strSQL,objConnection)
        objCommand.ExecuteNonQuery()

        objCommand = nothing
        objConnection.Close()
        objConnection = nothing

        'Redirect to confirm delete page
        Server.Transfer("categorymanager.aspx")
    
 End Sub


  'Count the number of recipes in the selected category
  Sub DisplayCount()

        Dim strSQL as string

        Dim getcatid as string
        getcatid = Request.QueryString("catid")

  strSQL = "SELECT Count(CAT_ID) FROM Recipes WHERE CAT_ID = " & Replace(getcatid, "'", "''")

       'Open database - connect to the database      
        objConnection = New OledbConnection(strConnection)
        objCommand = New OledbCommand(strSQL, objConnection)

        objCommand.Connection.Open()
lblrcdcount.Text = "Number of recipes belong to this category&nbsp; " & objCommand.ExecuteScalar() 
        objCommand.Connection.Close()

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
<div align="center">
<asp:Label cssClass="content6" id="lblrcdcount" runat="server" />
</div>
<br />
<div align="center"><span class="cred">Are you sure you want to delete this category?</span>
<br />
<span class="hlink">
<b>Note: </b>Deleting this category won't delete individual recipes associated with this category, you will have to find and delete each one of them.
</span></div>
<br />
<form runat="server">
                       <table width="40%" border="0" cellpadding="3" cellspacing="0" align="center">
            		                  <tr>
            			                 <td colspan=2  bgcolor="#6898d0">
            <span class="content3">Deleting Recipe Category</span>
            			                 </td>
            		                  </tr>
            		                  <tr>
          <td bgcolor="#f7f7f7" class="content2">Category Name:</td>   		
         <td bgcolor="#fbfbfb">
        <asp:TextBox runat="server" id="CategoryName" class="textbox" size="25" maxlenght="25" readonly="true" />
            			                 </td>
            		                  </tr>                 		                 
            		                  <tr>
            			       <td align=center colspan=2 bgcolor="#ffffff">
       <asp:Button runat="server" Text="Delete" id="deletebutton" class="submit" onclick="Delete_Category" />
            			             </td>
            		                  </tr>
            	                   </table>
        </form>
<br />
<div align="center"><asp:HyperLink cssClass="dt" tooltip="Back to Recipe Manager Main Page" runat="server" ID="approvallink" NavigateUrl="recipemanager.aspx">Recipe Manager Main Page</asp:HyperLink></div>
    </body>
</html>
