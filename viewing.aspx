<%@ Page Language="VB" Debug="true" %>
<%@ import Namespace="System.Data" %>
<%@ import Namespace="System.Data.Oledb" %>
<%@ import Namespace="System.Data.SqlClient" %>

<script runat="server">

   Sub Page_Load()
    
           Dim strSQL as string
           Dim totalcomments as integer
        
           'Call check user function - Check if user has started a session 
           Check_User()         
       
            'SQL display details and rating value
            strSQL = "SELECT *, (RATING/NO_RATES) AS Rates FROM Recipes WHERE id=" & Request.QueryString("id")
    
            'Connect to the database
            DataBase_Connect(strSQL)
    
            'Read data
            objDataReader.Read()
    
            lblname.text = objDataReader("Name")
            lblauthor.text = objDataReader("Author")
            lblhits.Text = objDataReader("Hits")
            lbldate.Text = objDataReader("Date")            
            Ingredients.text = objDataReader("Ingredients")
            Instructions.text = objDataReader("Instructions")
            
            'Close database connection
            DataBase_Disconnect()
    
  End Sub


</script>

<!--#include file="config.aspx"-->

<!--Powered By www.Ex-designz.net Recipe Cookbook ASP.NET version - Author: Dexter Zafra, Norwalk,CA-->

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title>Admin Viewing Recipe - www.ex-designz.net</title>
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
            <span class="content3">Viewing Recipe</span>
            			                 </td>
            		                  </tr>
            		                  <tr>
          <td bgcolor="#f7f7f7" class="content2">Name:</td>   		
         <td bgcolor="#fbfbfb">
<asp:Label runat="server" id="lblname" class="content2" />
            			                 </td>
            		                  </tr>
                     <tr>
                       <td bgcolor="#f7f7f7" class="content2">Author:</td>   		
                       <td bgcolor="#fbfbfb">
      <asp:Label runat="server" id="lblauthor" class="content2" />
            			 </td>        		                  
                                   </tr>

            		                  <tr>
          <td bgcolor="#f7f7f7" class="content2">Hits:</td>   		
         <td bgcolor="#fbfbfb">
<asp:Label runat="server" id="lblhits" class="content2" />
            			                 </td>
            		                  </tr>
                     <tr>
                       <td bgcolor="#f7f7f7" class="content2">Date:</td>   		
                       <td bgcolor="#fbfbfb">
                     <asp:Label runat="server" id="lbldate" class="content2" />
            			 </td>          		                  
                                    </tr>

            		                  <tr>
            			         <td valign="top" bgcolor="#f7f7f7" class="content2">Ingredients:</td>
            			            <td bgcolor="#fbfbfb">
 <asp:TextBox runat="server" id="Ingredients" Class="textbox" textmode="multiline" columns="70" rows="14" readonly />
            			                 </td>
            		                  </tr>
                                           <tr>
            			            <td valign="top" bgcolor="#f7f7f7" class="content2">Instructions:</td>  		
            			            <td bgcolor="#fbfbfb">
 <asp:TextBox runat="server" id="Instructions" Class="textbox" textmode="multiline" columns="70" rows="14" readonly />
<br />
<div style="text-align: left;" class="content2"><a href="editingunapproved.aspx?id=<%=Request.QueryString("id")%>" class="content2">Edit this recipe</a></div>
            			                 </td>
            		                  </tr>          		                  
            	                   </table>
<br />
<div style="text-align: center;" class="content2"><asp:HyperLink runat="server" NavigateUrl="JavaScript:onClick= window.close()" class="content2">Close Window</asp:HyperLink></div>
                                </td>
		                    </tr>
		               </table>
	               </td>
               </tr>
           </table>
    </form>

    </body>
</html>
