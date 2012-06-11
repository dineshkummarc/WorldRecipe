<%@ Page Language="VB" Debug="true" %>
<%@ import Namespace="System.Data" %>
<%@ import Namespace="System.Data.Oledb" %>
<%@ import Namespace="System.Data.SqlClient" %>

<script runat="server">

   Sub Page_Load()
    
            Dim strSQL as string
    
            Check_User()
    
    	    'Check which action were selected, edit a recipe or delete a recipe
            strSQL = "SELECT * FROM Recipes WHERE id=" & Request.QueryString("id") 
    
            DataBase_Connect(strSQL)   
            objDataReader.Read()

            'This will be the value to be populated into the textboxes
            Name.text = objDataReader("Name")
            Author.text = objDataReader("Author")
            Hits.text = objDataReader("Hits")
            Ingredients.text = objDataReader("Ingredients")
            Instructions.text = objDataReader("Instructions")
    
            DataBase_Disconnect()
    
 End Sub
    
   
    'Change any of recipes data, the name, ingredients, instructions, author 
   Sub Change_Recipes(sender As Object, e As System.EventArgs)

        Dim strSQL as string
    
        objConnection = New OledbConnection(strConnection)
        objConnection.Open()
    
        If request("CategoryName") = "0" AND request("CategoryID") = "0" Then
        strSQL = "update Recipes set Name='" & replace(request("Name"),"'","''")
        strSQL += "', Ingredients='" & replace(request("Ingredients"),"'","''")
        strSQL += "', Instructions='" & replace(request("Instructions"),"'","''")
        strSQL += "', Author='" & replace(request("Author"),"'","''")
        strSQL += "', Hits='" & replace(request("Hits"),"'","''")
        strSQL += "' where ID = " & request("id")

        Else
        
        strSQL = "update Recipes set Name='" & replace(request("Name"),"'","''")
        strSQL += "', CAT_ID='" & replace(request("CategoryName"),"'","''")
        strSQL += "', Category='" & replace(request("CategoryID"),"'","''")
        strSQL += "', Ingredients='" & replace(request("Ingredients"),"'","''")
        strSQL += "', Instructions='" & replace(request("Instructions"),"'","''")
        strSQL += "', Author='" & replace(request("Author"),"'","''")
        strSQL += "', Hits='" & replace(request("Hits"),"'","''")
        strSQL += "' where ID = " & request("id")

        End If

        objCommand = New OledbCommand(strSQL,objConnection)
        objCommand.ExecuteNonQuery()
    
        objCommand = nothing
        objConnection.Close()
        objConnection = nothing
        Server.Transfer("confirmupdate.aspx")
    
 End Sub
    
    'Delete the selected recipe
   Sub DeleteRecipes(sender As Object, e As System.EventArgs)
    
        Dim strSQL as string
    
        objConnection = New OledbConnection(strConnection)
        objConnection.Open()
    
        strSQL = "delete * from Recipes where ID = " & request("id")
        objCommand = New OledbCommand(strSQL,objConnection)
        objCommand.ExecuteNonQuery()
    
        objCommand = nothing
        objConnection.Close()
        objConnection = nothing

        'Redirect to confirm delete page
        Server.Transfer("confirmdel.aspx")
    
 End Sub
      
    'Event Back to recipe manager page
    Sub BackToManager(sender as object, e as System.EventArgs)
    
        Server.Transfer("recipemanager.aspx")
    
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
<div style="text-align: left; padding-left: 190px; margin-top: 12px;"><asp:HyperLink runat="server" NavigateUrl="recipemanager.aspx" class="content2">Back to Recipe Manager</asp:HyperLink></div>
           <table width=100% height=100%>
               <tr>
	               <td valign="middle">
                       <table width=40% border="0" cellpadding="0" cellspacing="0" align="center">
                            <tr>
                                <td bgcolor="#ffffff">
            	                   <table width="100%" border=0 cellpadding=3 cellspacing=1>
            		                  <tr>
            			                 <td colspan=2  bgcolor="#6898d0">
            <span class="content3">Edit / Delete Recipe</span>
            			                 </td>
            		                  </tr>
            		                  <tr>
          <td bgcolor="#f7f7f7" class="content2">Name:</td>   		
         <td bgcolor="#fbfbfb">
        <asp:TextBox runat="server" id="Name" class="textbox" size="30" maxlenght="30" />
            			                 </td>
            		                  </tr>
<tr>
          <td bgcolor="#f7f7f7" class="content2">Category:</td>   		
         <td bgcolor="#fbfbfb">
<span class="content8"><strong>Note:</strong> If you move the recipe to a different category, make sure you match the left field (Category Name) to the right field (Category ID), i.e. Barbque - Category ID = 1 has to match Category ID = 1. If you don't want to move it to a different category, don't do nothing, leave it as what is.</span>
<select name="CategoryName" size="1" id="CategoryName">
	<option value="0" selected>Category Name</option>
	<option value="1">Barbque - Category ID = 1</option>
	<option value="2">Beef - Category ID = 2</option>
	<option value="3">Breads - Category ID = 3</option>
	<option value="4">Cakes Desserts - Category ID = 4</option>
	<option value="5">Candy - Category ID = 5</option>
	<option value="6">Cassoroles - Category ID = 6</option>
	<option value="7">Dips - Category ID = 7</option>
	<option value="8">Drinks - Category ID = 8</option>
	<option value="9">Fish - Category ID = 9</option>
	<option value="10">Poultry - Category ID = 10</option>
	<option value="11">German - Category ID = 11</option>
	<option value="12">Lamb - Category ID = 12</option>
	<option value="13">Mexican - Category ID = 13</option>
	<option value="14">Oriental - Category ID = 14</option>
	<option value="15">PanCakes - Category ID = 15</option>
	<option value="16">Pies - Category ID = 16</option>
	<option value="17">Pork - Category ID = 17</option>
	<option value="18">Puddings - Category ID = 18</option>
	<option value="19">Russian - Category ID = 19</option>
	<option value="20">Salads - Category ID = 20</option>
	<option value="21">Sauces - Category ID = 21</option>
	<option value="22">SeaFoods - Category ID = 22</option>
	<option value="23">Soups - Category ID = 23</option>
	<option value="24">Syrups - Category ID = 24</option>
	<option value="25">Vegetables - Category ID = 25</option>
	<option value="26">Misc Unsorted - Category ID = 26</option>
	<option value="27">Afghan - Category ID = 27</option>
	<option value="28">Jewish - Category ID = 28</option>
	<option value="29">Korean - Category ID = 29</option>
	<option value="30">Japanese - Category ID = 30</option>
	<option value="31">Chinese - Category ID = 31</option>
	<option value="32">Filipino - Category ID = 32</option>
	<option value="33">Indian - Category ID = 33</option>
	<option value="34">Australian - Category ID = 34</option>
	<option value="35">African - Category ID = 35</option>
	<option value="36">American Indian - Category ID = 36</option>
	<option value="37">Irish Recipes - Category ID = 37</option>
	<option value="38">Jambalaya - Category ID = 38</option>
	<option value="39">Italian - Category ID = 39</option>
	<option value="40">Greek - Category ID = 40</option>
	<option value="41">Arabian - Category ID = 41</option>
	<option value="42">British - Category ID = 42</option>
	<option value="43">French - Category ID = 43</option>
	<option value="44">Thai - Category ID = 44</option>
	<option value="45">Dutch - Category ID = 45</option>
	<option value="46">Pakistan - Category ID = 46</option>
	<option value="47">Desserts - Category ID = 47</option>
	<option value="48">Cookies - Category ID = 48</option>
	<option value="49">Sandwich - Category ID = 49</option>

</select>

<select name="CategoryID" size="1" id="CategoryID">
	<option value="0" selected>Category ID</option>
	<option value="Barbque">1</option>
	<option value="Beef">2</option>
	<option value="Breads">3</option>
	<option value="Cakes Desserts">4</option>
	<option value="Candy">5</option>
	<option value="Cassoroles">6</option>
	<option value="Dips">7</option>
	<option value="Drinks">8</option>
	<option value="Fish">9</option>
	<option value="Poultry">10</option>
	<option value="German">11</option>
	<option value="Lamb">12</option>
	<option value="Mexican">13</option>
	<option value="Oriental">14</option>
	<option value="PanCakes">15</option>
	<option value="Pies">16</option>
	<option value="Pork">17</option>
	<option value="Puddings">18</option>
	<option value="Russian">19</option>
	<option value="Salads">20</option>
	<option value="Sauces">21</option>
	<option value="SeaFoods">22</option>
	<option value="Soups">23</option>
	<option value="Syrups">24</option>
	<option value="Vegetables">25</option>
	<option value="Misc Unsorted">26</option>
	<option value="Afghan">27</option>
	<option value="Jewish">28</option>
	<option value="Korean">29</option>
	<option value="Japanese">30</option>
	<option value="Chinese">31</option>
	<option value="Filipino">32</option>
	<option value="Indian">33</option>
	<option value="Australian">34</option>
	<option value="African">35</option>
	<option value="American Indian">36</option>
	<option value="Irish Recipes">37</option>
	<option value="Jambalaya">38</option>
	<option value="Italian">39</option>
	<option value="Greek">40</option>
	<option value="Arabian">41</option>
	<option value="British">42</option>
	<option value="French">43</option>
	<option value="Thai">44</option>
	<option value="Dutch">45</option>
	<option value="Pakistan">46</option>
	<option value="Desserts">47</option>
	<option value="Cookies">48</option>
	<option value="Sandwich">49</option>

</select>

            			                 </td>
            		                  </tr>
                     <tr>
                       <td bgcolor="#f7f7f7" class="content2">Author:</td>   		
                       <td bgcolor="#fbfbfb">
                     <asp:TextBox runat="server" id="Author" class="textbox" size="25" maxlenght="25" />
            			 </td>
            		                  </tr>
<tr>
          <td bgcolor="#f7f7f7" class="content2">Hits:</td>   		
         <td bgcolor="#fbfbfb">
<asp:TextBox runat="server" id="Hits" class="textbox" size="6" maxlenght="6" />
            			                 </td>
            		                  </tr>
            		                  <tr>
            			         <td valign="top" bgcolor="#f7f7f7" class="content2">Ingredients:</td>
            			            <td bgcolor="#fbfbfb">
            <asp:TextBox runat="server" id="Ingredients" Class="textbox" textmode="multiline" columns="70" rows="14" />
            			                 </td>
            		                  </tr>
                                           <tr>
            			            <td valign="top" bgcolor="#f7f7f7" class="content2">Instructions:</td>  		
            			            <td bgcolor="#fbfbfb">
            <asp:TextBox runat="server" id="Instructions" Class="textbox" textmode="multiline" columns="70" rows="14" />
            			                 </td>
            		                  </tr>
            		                  <tr>
            			       <td align=center colspan=2 bgcolor="#ffffff">
       <asp:Button runat="server" Text="Update" id="updatebutton" class="submit" onclick="Change_Recipes"/>
       <asp:Button runat="server" Text="Delete" id="deletebutton" class="submit" onclick="DeleteRecipes"/>
       <asp:Button runat="server" Text="Cancel" id="cancelbutton" class="submit" onclick="BackToManager"/>
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
