<%@ Page Language="VB" Debug="true" %>
<%@ import Namespace="System.Data" %>
<%@ import Namespace="System.Data.Oledb" %>
<%@ import Namespace="System.Data.SqlClient" %>

<script runat="server">

   Sub Page_Load()
    

            Dim strSQL as string
    
            Check_User()
    
    	    'Check which action were selected, edit a recipe or delete a recipe
            strSQL = "SELECT * FROM COMMENTS_RECIPE WHERE id=" & Request.QueryString("id") 
    
            DataBase_Connect(strSQL)   
            objDataReader.Read()

            'This will be the value to be populated into the textboxes
            Author.text = objDataReader("AUTHOR")
            Email.text = objDataReader("EMAIL")
            Comments.text = objDataReader("COMMENTS")
    
            DataBase_Disconnect()
    
 End Sub
    
   

    'Change any of recipes data, the name, ingredients, instructions, author 
   Sub Update_Comments(sender As Object, e As System.EventArgs)

        Dim strSQL as string
    
        objConnection = New OledbConnection(strConnection)
        objConnection.Open()
        
        strSQL = "update COMMENTS_RECIPE set AUTHOR='" & replace(request("Author"),"'","''")
        strSQL += "', EMAIL='" & replace(request("Email"),"'","''")
        strSQL += "', COMMENTS='" & replace(request("Comments"),"'","''")
        strSQL += "' where ID = " & request("id")

        objCommand = New OledbCommand(strSQL,objConnection)
        objCommand.ExecuteNonQuery()
    
        objCommand = nothing
        objConnection.Close()
        objConnection = nothing
        Server.Transfer("commentsmanager.aspx")
    
 End Sub
    


   'Delete the selected comments
   Sub Delete_Comments(sender As Object, e As System.EventArgs)
    
        Dim strSQL as string
    
        objConnection = New OledbConnection(strConnection)
        objConnection.Open()
    
        strSQL = "delete * from COMMENTS_RECIPE where ID = " & request("id")
        objCommand = New OledbCommand(strSQL,objConnection)
        objCommand.ExecuteNonQuery()

        objCommand = nothing
        objConnection.Close()
        objConnection = nothing
    

       Dim strSQL2 as string

       objConnection = New OledbConnection(strConnection)
       objConnection.Open()

    'Subtract 1 to the total comments when comment is deleted             
     strSQL2 = "Update Recipes set TOTAL_COMMENTS = 0  where id=" & Request.QueryString("id")

        objCommand = New OledbCommand(strSQL2,objConnection)
        objCommand.ExecuteNonQuery()

        objCommand = nothing
        objConnection.Close()
        objConnection = nothing

        'Redirect to confirm delete page
        Server.Transfer("commentsmanager.aspx")
    
 End Sub
      
    'Event Back to recipe manager page
    Sub BackToManager(sender as object, e as System.EventArgs)
    
        Server.Transfer("commentsmanager.aspx")
    
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
<form runat="server">
           <table width=100% height=100%>
               <tr>
	               <td valign="middle">
                       <table width=40% border="0" cellpadding="0" cellspacing="0" align="center">
                            <tr>
                                <td bgcolor="#ffffff">
            	                   <table width="100%" border=0 cellpadding=3 cellspacing=1>
            		                  <tr>
            			                 <td colspan=2  bgcolor="#6898d0">
            <span class="content3">Edit / Delete Comments</span>
            			                 </td>
            		                  </tr>
            		                  <tr>
          <td bgcolor="#f7f7f7" class="content2">Author:</td>   		
         <td bgcolor="#fbfbfb">
        <asp:TextBox runat="server" id="Author" class="textbox" size="25" maxlenght="25" />
            			                 </td>
            		                  </tr>
                     <tr>
                       <td bgcolor="#f7f7f7" class="content2">Email:</td>   		
                       <td bgcolor="#fbfbfb">
                     <asp:TextBox runat="server" id="Email" class="textbox" size="35" maxlenght="35" />
            			 </td>
            		                  </tr>
            		                 <tr>
            			        <td valign="top" bgcolor="#f7f7f7" class="content2">Comments:</td>
            			            <td bgcolor="#fbfbfb">
            <asp:TextBox runat="server" id="Comments" Class="textbox" textmode="multiline" columns="70" rows="14" />
            			                 </td>
            		                  </tr>
            		                  <tr>
            			       <td align=center colspan=2 bgcolor="#ffffff">
       <asp:Button runat="server" Text="Update" id="updatebutton" class="submit" onclick="Update_Comments" />
       <asp:Button runat="server" Text="Delete" id="deletebutton" class="submit" onclick="Delete_Comments" />
       <asp:Button runat="server" Text="Cancel" id="cancelbutton" class="submit" onclick="BackToManager" />
            			             </td>
            		                  </tr>
            	                   </table>
                                </td>
		                    </tr>
		               </table>
	               </td>
               </tr>
           </table>
        </form>
    </body>
</html>
