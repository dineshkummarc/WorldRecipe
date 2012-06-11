<%@ Page Language="VB" Debug="true" %>
<%@ import Namespace="System.Data" %>
<%@ import Namespace="System.Data.Oledb" %>
<%@ Import Namespace="System.Web.Mail" %>

<script runat="server">

 'Handle page load event
 Sub Page_Load(Sender As Object, E As EventArgs)
       
           'Call add hits sub routine
           AddHits()          

           'Call the DisplayRating sub routine
           DisplayRating()

           'Call NewestRecipes Sub - Display 15 newest recipes
           NewestRecipes()

           'Call MostPopular Sub - Display 15 most popular recipes
           MostPopular()
           
           'Call risplay recipe name A to Z
           DisplayLetter()

           'Call display comments
           DisplayComments()

           GetCatCategory()

           'Call Random Recipe
           RandomRecipeNumber()
           RandomRecipe()


           lblletter.Text = "Recipe Name List A-Z:"
           lblsearch.Text = "Search recipe by name,author or country of origin i.e.(Filipino,chinese)"
           lblNamedis.text = "Name:"
           lblcategorydis.text = "Category:"
           lblauthordis.text = "Author:"
           lbldatedis.text = "Date:"
           lblhitsdis.text = "Hits:"
           lblratingdis.text = "Rating:"
           lblyourratingdis.text = "Please Rate This Recipe:&nbsp;&nbsp;"
           lblingredientsdis.text = "Ingredients:"
           lblinstructionsdis.text = "Instructions:"
           lblheadersearch.text = "Search Recipe"
           lblheadermostpopular.text = "15 Most Popular"
           lblheadernewest.text = "15 New Recipes"
           lblcommentsdis.text = "Comments:"
           lblallfieldsrequired.text = "All fields are required!"
           lblyournamedis.text = "Your Name:"
           lblyouremaildis.text = "Email:"
           lblheaderrandom.text = "Featured Recipe"
           lblPagename.text = "Recipe Detail"
           lblCatNameHeader.text = "Choose Category" 
              

 'SQL display details and rating value
 strSQL = "SELECT * FROM Recipes WHERE LINK_APPROVED = 1 AND id=" & Request.QueryString("id")
    
            Dim objDataReader as OledbDataReader
            
            'Call Open database - connect to the database
            DBconnect()

            objConnection.Open()
            objDataReader  = objCommand.ExecuteReader()
    
            'Read data
            objDataReader.Read()
    
            lblviewing.text = "Viewing Recipe"
            lblname.text = objDataReader("Name")
            lblname2.text = "Write a Comment for&nbsp;" & objDataReader("Name") & "&nbsp;recipe"
            lblauthor.text = objDataReader("Author")
            lblhits.Text = objDataReader("Hits")
            lblcategorytop.Text = "Other&nbsp;" & objDataReader("Category") & "&nbsp;recipes you might be interested"
            lbldate.Text = objDataReader("Date")            
            lblIngredients.text = Replace(objDataReader("Ingredients"), Chr(13), "<br>")
            lblInstructions.text = Replace(objDataReader("Instructions"), Chr(13), "<br>")
            strRName = objDataReader("Name")
            strCName = objDataReader("Category")


            HyperLink2.NavigateUrl = "category.aspx?&catid=" & objDataReader("CAT_ID")
            HyperLink2.Tooltip = "Go to " & objDataReader("Category") & " recipe category"
            HyperLink2.Text = objDataReader("Category")
            HyperLink3.NavigateUrl = "category.aspx?&catid=" & objDataReader("CAT_ID")
            HyperLink3.Tooltip = "Go to " & objDataReader("Category") & " recipe category"
            HyperLink3.Text = objDataReader("Category")
            HyperLink4.NavigateUrl = "index.aspx"
            HyperLink4.Text = "Recipe Home"
            HyperLink4.ToolTip = "Back to recipe homepage"
 

            Dim totalcomments as integer

            'Get total comments value from the total comments field           
            totalcomments = objDataReader("TOTAL_COMMENTS")

       'Check the total comments value if it is greater or equal than one
       If totalcomments >= 1 Then

            'Display the total comments value, and enabled the hyperlink
            ReadComments.text = "There are:&nbsp;" & "(" & objDataReader("TOTAL_COMMENTS") & ")" & "&nbsp;comments"

            Elseif totalcomments = 0 Then

            'Display the total comments value, and disabled the hyperlink
            ReadComments.text = "There are no comments:&nbsp;" & "(" & objDataReader("TOTAL_COMMENTS") & ")"

       End If
 
                   
         Dim Popular as string
         Popular = objDataReader("CAT_ID")

strSQL = "SELECT Top 10 ID, CAT_ID,Name,HITS FROM Recipes WHERE CAT_ID =" & Replace(Popular, "'", "''") & "  ORDER BY ID ASC"

         'Call Open database - connect to the database
         DBconnect()

         Dim RecipeAdapter as New OledbDataAdapter(objCommand)
         Dim dts as New DataSet()
         RecipeAdapter.Fill(dts, "Name")

         RecipeTop.DataSource = dts.Tables("Name").DefaultView
         RecipeTop.DataBind()
 
      
        'Close database connection for the objDataReader
         objDataReader.Close()
         objConnection.Close() 
          
  End Sub

 

'Display Category
 Sub GetCatCategory()
         
         strSQL = "SELECT * FROM RECIPE_CAT Order By CAT_TYPE ASC"

         'Call Open database - connect to the database
         DBconnect()

         Dim RecipeAdapter as New OledbDataAdapter(objCommand)
         Dim dts as New DataSet()
         RecipeAdapter.Fill(dts, "CAT_ID")

         CategoryName.DataSource = dts.Tables("CAT_ID").DefaultView
         CategoryName.DataBind()


 End Sub



 'Display Comments
 Sub DisplayComments()
        
 strSQL = "SELECT * From COMMENTS_RECIPE Where ID=" & Request.QueryString("id") & " Order By Date Desc"

         'Call Open database - connect to the database
         DBconnect()

         Dim RecipeAdapter as New OledbDataAdapter(objCommand)
         Dim dts as New DataSet()
         RecipeAdapter.Fill(dts, "ID")

         RecComments.DataSource = dts         
         RecComments.DataBind() 

  End Sub


'Display recipe name A to Z
 Sub DisplayLetter()
         
strSQL = "SELECT * FROM Recipe_Letter Order By Letter ASC"

         'Call Open database - connect to the database
         DBconnect()

         Dim RecipeAdapter as New OledbDataAdapter(objCommand)
         Dim dts as New DataSet()
         RecipeAdapter.Fill(dts, "ID")

         RecipeLetter.DataSource = dts.Tables("ID").DefaultView
         RecipeLetter.DataBind()


 End Sub



  'Display 15 newest recipes
  Sub NewestRecipes()
      
strSQL = "SELECT Top 15 ID,Name,HITS,Category FROM Recipes Where LINK_APPROVED = 1 Order By Date DESC"

        'Call Open database - connect to the database
         DBconnect()

         Dim RecipeAdapter as New OledbDataAdapter(objCommand)
         Dim dts as New DataSet()
         RecipeAdapter.Fill(dts, "ID")

         RecipeNew.DataSource = dts.Tables("ID").DefaultView
         RecipeNew.DataBind()

 End Sub



  'Display 15 most popular recipes
  Sub MostPopular()
          
strSQL = "SELECT Top 15 ID,Name,HITS,Category FROM Recipes Where LINK_APPROVED = 1 Order By HITS DESC"

         'Call Open database - connect to the database
         DBconnect()

         Dim RecipeAdapter as New OledbDataAdapter(objCommand)
         Dim dts as New DataSet()
         RecipeAdapter.Fill(dts, "ID")

         RecipeTopMost.DataSource = dts.Tables("ID").DefaultView
         RecipeTopMost.DataBind()

 End Sub



  'Page level error handling - If page encounter an error, then redirect to the custom error page
  'Protected Overrides Sub OnError(ByVal e As System.EventArgs)

     'Server.Transfer("error.aspx")

  'End Sub



  'Increment hits by 1 every time a page load
  Sub AddHits()               

            strSQL = "Update Recipes set HITS = HITS + 1  where id=" & Request.QueryString("id")

            'Call Open database - connect to the database
            DBconnect()

            objConnection.Open()
            objCommand.ExecuteNonQuery()
    
            'Close the db connection and free up memory
            DBclose()

  End Sub

 
 
 'Display the star rating
 Sub DisplayRating()
 
 strSQL = "SELECT *, (RATING/NO_RATES) AS Rates FROM Recipes WHERE LINK_APPROVED = 1 AND ID =" & Request.QueryString("id")
          
         'Call Open database - connect to the database
         DBconnect()

         Dim RecipeAdapter as New OledbDataAdapter(objCommand)
         Dim dts as New DataSet()
         RecipeAdapter.Fill(dts, "ID")

         RecipeCat.DataSource = dts.Tables("ID").DefaultView
         RecipeCat.DataBind()

 End Sub



 'Insert users rating to the database
 Sub Get_Rating(sender As Object, e As System.EventArgs)

     If Page.IsPostBack Then
        
'SQL insert rating   
strSQL = "Update Recipes  SET RATING = RATING + " & Replace(Request("RateMe"),"'","''") & ", NO_RATES = NO_RATES + 1 WHERE ID =" & Request.QueryString("id")
        
    
           'Call Open database - connect to the database
           DBconnect()

           objConnection.Open()
           objCommand.ExecuteNonQuery()
    
           'Close the db connection and free up memory
           DBclose()
           
           'Redirect to previous page after rating is complete
           Dim urlredirect as string
           urlredirect = "recipedetail.aspx?&id=" & Request.QueryString("id")
           Server.Transfer(urlredirect)


    End if
    
 End Sub



 'Insert comment to the database
 Sub Add_Comment(sender As Object, e As System.EventArgs)
    
     'Do the validation of the add comment fields before inserting to the database
    
     If Page.IsPostBack Then

     If Page.IsValid then
            
           'SQL insert comment
           strSQL = "insert into COMMENTS_RECIPE (ID,AUTHOR,EMAIL,COMMENTS) values ('" &        replace(request("id"),"'","''")
           strSQL += "','" & replace(request("AUTHOR"),"'","''")
           strSQL += "','" & replace(request("EMAIL"),"'","''") 
           strSQL += "','" & replace(request("COMMENTS"),"'","''") & "')"

           'Call Open database - connect to the database
           DBconnect()
    
           objConnection.Open()
           objCommand.ExecuteNonQuery()
    
           'Close the db connection and free up memory
           DBclose()


          'This part is the email notification when someone write a comment  
          Dim strBody As String

         strBody = "Hello Webmaster, Someone has wrote a recipe comment:" _
	 & vbCrLf & vbCrLf _
         & "http://" & Request.ServerVariables("HTTP_HOST") _
         & Request.ServerVariables("URL") & "?id=" & Request.QueryString("id") & vbCrLf

         Dim mailnotify As SmtpMail
         Dim NotifyEmail As New MailMessage()

         'Email notification - Change the email (extremedexter_z2001@yahoo.com) 
         'to your domainemail or any email address you have.
         NotifyEmail.To = "extremedexter_z2001@yahoo.com"
         NotifyEmail.From = "recipecommentnotify@myasp-net.com"
         NotifyEmail.Subject = "Myasp-net.net Recipe Comment Notification"
         NotifyEmail.Body = strBody
         mailnotify.SmtpServer = "localhost" 
         mailnotify.Send(NotifyEmail)


         'Call the NumberComments - Increment 1 the total of comments
         NumberComments()


    End if

  End If
    
 End Sub

 

 'Increment 1 to the total_comments 
 Sub NumberComments()
            
       'SQL increment 1 total comments  
       strSQL = "Update Recipes SET TOTAL_COMMENTS = TOTAL_COMMENTS + 1 where ID =" & Request.QueryString("id")
            
          'Call Open database - connect to the database
          DBconnect()

          objConnection.Open()
          objCommand.ExecuteNonQuery()
          
          'Close the db connection and free up memory
          DBclose()

          'Redirect back to previous page upon success adding comment
          Dim urlredirect2 as string
          urlredirect2 = "recipedetail.aspx?&id=" & Request.QueryString("id")
          Server.Transfer(urlredirect2)

 End Sub



'Pulls a Random number for selecting a random recipe
  Sub RandomRecipeNumber()

        'It connects to database
        strSQL = "SELECT ID FROM Recipes"

        Dim objDataReader as OledbDataReader
            
       'Call Open database - connect to the database
        DBconnect()

        objConnection.Open()
        objDataReader  = objCommand.ExecuteReader()

        'Counts how many records are in the database
        Dim iRecordNumber = 0
        do while objDataReader.Read()=True
            iRecordNumber += 1
        loop

        objDataReader.Close()
        objConnection.Close()

        'Here's where random number is generated
        Randomize()
        do
            iRandomRecipe = (Int(RND() * iRecordNumber))
        loop until iRandomRecipe <> 0


  End Sub



  'Pulls aand dsiplay random recipe records
  Sub RandomRecipe()

        strSQL = "SELECT ID,CAT_ID,Category,Name,Author,Date,HITS,RATING,NO_RATES, (RATING/NO_RATES) AS Rates FROM Recipes"

         Dim objDataReader as OledbDataReader
            
        'Call Open database - connect to the database
        DBconnect()

        objConnection.Open()
        objDataReader  = objCommand.ExecuteReader()

        Dim i = 0

        'Go until a random position
        do while i<>iRandomRecipe
            objDataReader.Read()
            i += 1
        loop

        Dim strRanRating as Double

        'Display recipe
        lblRating2.Text = "Rating:"
        lblRancategory.text = "Category:" 
        lblranhitsdis.text = "Hits:"
        lblranhits.text = objDataReader("Hits")
        strRanRating = FormatNumber(objDataReader("Rates"), 1,  -2, -2, -2)
        lblranrating.Text = "(" & strRanRating & ")"
        strRatingimg = FormatNumber(objDataReader("Rates"), 1,  -2, -2, -2)

        LinkRanName.NavigateUrl = "recipedetail.aspx?id=" & objDataReader("ID")
        LinkRanName.Text = objDataReader("Name")
        LinkRanName.ToolTip = "View" & " - " & objDataReader("Name") & " - " & "recipe"
        LinkRanCat.NavigateUrl = "category.aspx?catid=" & objDataReader("CAT_ID")
        LinkRanCat.Text = objDataReader("Category")
        LinkRanCat.ToolTip = "Go to" & " - " & objDataReader("Category") & " - " & "&category"

        objDataReader.Close()
        objConnection.Close()

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
    Public strRatingimg as Integer
    Public iRandomRecipe as integer
    Public strRName as string
    Public strCName as string


</script>

<!--#include file="inc_databasepath.aspx"--> 
<!--#include file="inc_headrecipedetail.aspx"--> 

<form runat="server" style="margin-top: 0px; margin-bottom: 0px;">
   <table border="0" cellpadding="0" cellspacing="0" width="100%">
     <tr>
      <td width="15%" valign="top" align="left">
     <!--#include file="inc_navmenu.aspx"-->
    <span class="content3">br</span>
<!--Begin Display Category Menu-->
  <div class="div8">
  <div class="div9"><asp:Label cssClass="content3" runat="server" id="lblCatNameHeader" /></div>
 <div class="div6">
<asp:DataList cssClass="hlink" id="CategoryName" runat="server">
   <ItemTemplate>
<div style="padding: 0;">
<span class="bluearrow3">�</span>
<a class="dt" title="Go to <%# DataBinder.Eval(Container.DataItem, "CAT_TYPE") %> category" href='<%# DataBinder.Eval(Container.DataItem, "CAT_ID", "category.aspx?catid={0}") %>'><%# DataBinder.Eval(Container.DataItem, "CAT_TYPE") %></a>
</div>
   </ItemTemplate>
  </asp:DataList>
 </div>
</div>
<!--End Display Category Menu-->
  </td>
    <td width="70%" valign="top">
      <table width=100% height=100%>
        <tr>
	  <td valign="top">
            <table width=100% border="0" cellpadding="0" cellspacing="0" align="center">
              <tr>
                <td bgcolor="#ffffff">
            	   <table width="100%" border="0" cellpadding="3" cellspacing="1">
                    <tr>
            	     <td colspan="2" bgcolor="#ffffff">
                       <div style="text-align: left; padding-left: 0;  margin-bottom: 10px;">
                        <asp:HyperLink id="HyperLink4" cssClass="dtcat" runat="server" />
                       <span class="bluearrow">�</span>
                     <asp:HyperLink id="HyperLink3" cssClass="dtcat" runat="server" />
                   <span class="bluearrow">�</span>&nbsp;<asp:Label cssClass="content10" runat="server" id="lblPagename" />
                     </div>
<div style="padding: 2px; text-align: center; border: 1px dashed #e5e5e5; margin-bottom: 10px; margin-left: 26px; margin-right: 26px;">
       <asp:Label cssClass="corange" runat="server" id="lblletter" />
    <asp:DataList cssClass="hlink" id="RecipeLetter" RepeatColumns="26" runat="server">
   <ItemTemplate>
 <div style="padding: 4px;">
<a class="letter" title="View all recipes starting with letter <%# DataBinder.Eval(Container.DataItem, "Letter") %>" href='<%# DataBinder.Eval(Container.DataItem, "Letter", "letter.aspx?l={0}") %>'>
<%# DataBinder.Eval(Container.DataItem, "Letter") %></a>
</div>
    </ItemTemplate>
  </asp:DataList>
</div>
     </td>
         </tr>
             <tr>
             <td bgcolor="#e6edf7">
            <table border="0" cellpadding="0" cellspacing="0" width="100%">
  <tr>
    <td width="30%" height="20" bgcolor="#e6edf7">
    <asp:Label cssClass="content6" runat="server" id="lblviewing" />
    </td>
    <td width="70%" align="right" height="20" bgcolor="#e6edf7">
    <div style="margin-right: 10px;">
<img src="images/favstar.gif" align="middle" alt="Add <%=strRName%> recipe to favorite"> 
<a class="dt" title="Add <%=strRName%> recipe to favorite" href="JavaScript:window.external.AddFavorite(location.href, document.title)">Add to fav</a>&nbsp;&nbsp;
<img src="images/email.gif" align="middle" alt="Email <%=strRName%> recipe to friend"> 
<a class="dt" title="Email <%=strRName%> recipe to friend" href="javascript:openWindow('emailrecipe.aspx?url=http://www.myasp-net.com/recipedetail.aspx?id=<%=Request.QueryString("id")%>&amp;n=<%=strRName%>&amp;c=<%=strCName%>')">Email Recipe</a>&nbsp;&nbsp;
<img src="images/print.gif" align="middle" alt="Print <%=strRName%> recipe"> 
<a class="dt" title="Print <%=strRName%> recipe" href="javascript:Start('print.aspx?id=<%=Request.QueryString("id")%>')">Print Recipe</a>
</div>
    </td>
  </tr>
</table>
       </td>
      </tr>
   <tr>  	
   <td bgcolor="#fcfcfc">
<asp:Label cssClass="content2" runat="server" id="lblNamedis" /> 
<asp:Label cssClass="cmaron" runat="server" id="lblname" />
      </td>
        </tr>
          <tr>
           <td bgcolor="#fcfcfc">
<asp:Label cssClass="content2" runat="server" id="lblcategorydis" /> 
<asp:HyperLink id="HyperLink2" cssClass="dt" runat="server" />
</td>
 </tr>
   <tr>
     <td bgcolor="#fcfcfc">
<asp:Label cssClass="content2" runat="server" id="lblauthordis"/> 
<asp:Label runat="server" id="lblauthor" class="content2" />
   </td>        		                  
</tr>
   <tr>
      <td bgcolor="#fcfcfc">
<asp:Label cssClass="content2" runat="server" id="lbldatedis" />
<asp:Label runat="server" id="lbldate" class="content2" />
    </td>          		                  
</tr>
     <tr>
       <td bgcolor="#fcfcfc">
<asp:Label cssClass="content2" runat="server" id="lblhitsdis" />
<asp:Label runat="server" id="lblhits" class="content2" />
   </td>
</tr>
  <tr>  		
    <td bgcolor="#fcfcfc">
<asp:Label cssClass="content2" runat="server" id="lblratingdis" />
<asp:DataList cssClass="hlink" id="RecipeCat" RepeatColumns="1" runat="server">
   <ItemTemplate>
<img src="images/<%# FormatNumber((DataBinder.Eval(Container.DataItem, "Rates")), 1,  -2, -2, -2) %>.gif" style="vertical-align: middle;" alt="Rating <%# FormatNumber((DataBinder.Eval(Container.DataItem, "Rates")), 1, -2, -2, -2) %>">&nbsp;<span class="content2">( <%# FormatNumber((DataBinder.Eval(Container.DataItem, "Rates")), 1, -2, -2, -2) %> ) by <%# DataBinder.Eval(Container.DataItem, "NO_RATES") %> users</span>
    </ItemTemplate>
  </asp:DataList>
     </td>
   </tr>
<tr>
   <td width="100%" bgcolor="#ffffff">
    <div style="border: solid 1px #E1EDFF; padding-top: 5px; padding-bottom: 15px; padding-left: 8px; width: auto; margin-top: 10px;">
<asp:Label cssClass="content6" runat="server" id="lblingredientsdis" />
  <br />
  <asp:Label cssClass="drecipe" ID="lblIngredients" runat="server" />
 </div>
</td>
 </tr>
  <tr>
   <td width="100%" bgcolor="#ffffff">
    <div style="border: solid 1px #E1EDFF; padding-top: 5px; padding-bottom: 15px; padding-left: 8px; width: auto; margin-top: 10px;">
<asp:Label cssClass="content6" runat="server" id="lblinstructionsdis" />
  <br />
  <asp:Label cssClass="drecipe" ID="lblInstructions" runat="server" />
</div>
  </td>
</tr>
     <tr>
       <td>
          <div style="margin-left: 10px;">
   <asp:Label cssClass="content2" BackColor="#F4F9FF" runat="server" id="lblyourratingdis" />
   <asp:Dropdownlist cssClass="cselect" runat="server" ID="RateMe" AutoPostBack="True" OnSelectedIndexChanged="Get_Rating">
          <asp:ListItem Text="Your Rating" Value="5" Selected="True" />
          <asp:ListItem Text="Excellent - 5 stars" Value="5" />
          <asp:ListItem Text="Very Good - 4 stars" Value="4" />
          <asp:ListItem Text="Interesting - 3 stars" Value="3" />
          <asp:ListItem Text="Fair - 2 stars" Value="2" />
          <asp:ListItem Text="Not sure - 1 star" Value="1" />
    </asp:DropDownList>
      </div>
     </td>
   </tr>          		                  
  </table>
 </td>
</tr>
</table>
  </td>
 </tr>
</table>
<br />
<table border="0" align="center" cellpadding="0" cellspacing="0" width="100%">
  <tr>
    <td width="100%">
<div style="margin-left: 14px; margin-bottom: 22px;">
<asp:Label cssClass="content4" runat="server" id="lblcategorytop" />
   <asp:DataList cssClass="hlink" id="RecipeTop" RepeatColumns="1" runat="server">
   <ItemTemplate>
<span class="bluearrow2">�</span>&nbsp;
<a class="dt" title="View <%# DataBinder.Eval(Container.DataItem, "Name") %> recipe" href='<%# DataBinder.Eval(Container.DataItem, "ID", "recipedetail.aspx?id={0}") %>'>
<%# DataBinder.Eval(Container.DataItem, "Name") %></a>
      </ItemTemplate>
  </asp:DataList>
</div>
   </td>
  </tr>
</table>
<!--Begin Display Comments-->
<table border="0" cellpadding="0" cellspacing="0" align="center" width="97%">
  <tr>
    <td width="100%" height="18" BgColor="#F4F9FF"><asp:Label id="ReadComments" cssClass="content6" runat="server" /></td>
  </tr>
  <tr>
    <td width="100%">
<asp:DataList width="100%" id="RecComments" RepeatColumns="1" runat="server">
      <ItemTemplate>
    <div class="divwrap2">
<div class="divbd2">
<b>Author:</b>&nbsp;<%# DataBinder.Eval(Container.DataItem, "AUTHOR") %>
<br />
<b>Email:</b>&nbsp;<%# DataBinder.Eval(Container.DataItem, "EMAIL") %>
<br />
<b>Date:</b>&nbsp; <%# FormatDateTime(DataBinder.Eval(Container.DataItem, "DATE"),vbShortDate) %>
<br />
<b>Comment:</b>
<br /> 
<%# Replace(DataBinder.Eval(Container.DataItem, "COMMENTS"), Chr(13), "<br>") %>
     </div>
   </div>
 </ItemTemplate>
</asp:DataList>
</td>
  </tr>
</table>
<!--End Display Comments-->
<table border="0" align="center" cellpadding="2" cellspacing="2" width="100%">
  <tr>
    <td width="100%">
<div style="border: solid 1px #E1EDFF; width: auto; margin-top: 20px;">
<table border="0" align="center" cellpadding="2" cellspacing="2" width="60%">
  <tr>
    <td width="100%" colspan="2">
</td>
  </tr>
  <tr>
    <td width="100%" colspan="2"><asp:Label cssClass="content4" runat="server" id="lblname2" />
<br />
<br />
<asp:Label cssClass="content5" runat="server" id="lblallfieldsrequired" /></td>
  </tr>
  <tr>
    <td width="21%" class="content2"><asp:Label cssClass="content2" runat="server" id="lblyournamedis" /></td>
    <td width="79%"><asp:TextBox ID="AUTHOR" Class="textbox" runat="server" size="20" maxlenght="20" />
<asp:RequiredFieldValidator runat="server"
      id="reqName" ControlToValidate="AUTHOR"
      cssClass="cred2"
      ErrorMessage = "Enter your name!"
      display="Dynamic" />
</td>
  </tr>
  <tr>
    <td width="21%" class="content2"><asp:Label cssClass="content2" runat="server" id="lblyouremaildis" /></td>
    <td width="79%"><asp:TextBox ID="EMAIL" Class="textbox" runat="server" size="25" maxlenght="25" />
 <asp:RequiredFieldValidator runat="server"
      id="reqEmail" ControlToValidate="EMAIL"
      cssClass="cred2"
      ErrorMessage = "Enter your email!"
      display="Dynamic">
 </asp:RequiredFieldValidator>
 <asp:RegularExpressionValidator id="RegularExpressionValidator1" runat="server"
            ControlToValidate="EMAIL"
            ValidationExpression="^[\w-]+@[\w-]+\.(com|net|org|edu|mil)$"
            Display="Static"
            cssClass="cred2">
 Enter a valid e-mail
 </asp:RegularExpressionValidator>
</td>
  </tr>
  <tr>
    <td width="21%" valign="top" class="content2"><asp:Label cssClass="content2" runat="server" id="lblcommentsdis" /></td>
    <td width="79%"><asp:TextBox ID="COMMENTS" Class="textbox" textmode="multiline" columns="45" rows="7"  runat="server" />
<br />
<asp:RequiredFieldValidator runat="server"
      id="reqComments" ControlToValidate="COMMENTS"
      cssClass="cred2"
      ErrorMessage = "Enter a comment!"
      display="Dynamic" />
<br />
<input type="hidden" value="<%=Request.QueryString("id")%>" ID="ID" name="ID">
<asp:Button runat="server" Text="Submit" id="AddComments" class="submit" onclick="Add_Comment"/>
      </td>
    </tr>
   </table>
  </td>
 </tr>
</table>
</div>
</form>
 </td>
   <td width="15%" valign="top" valign="top" align="left" class="content2">
<!--Begin Random Recipe-->
    <div class="div8">
<div class="div9"><asp:Label cssClass="content3" runat="server" id="lblheaderrandom" /></div>
 <div class="div6">
<span class="bluearrow2">�</span>
<asp:HyperLink id="LinkRanName" cssClass="dtcat" runat="server" />
<br />
<asp:Label cssClass="content8" runat="server" id="lblRancategory" /> <asp:HyperLink id="LinkRanCat" cssClass="dt2" runat="server" />
<br />
<asp:Label cssClass="content8" runat="server" id="lblranhitsdis" /> <asp:Label cssClass="content8" runat="server" id="lblranhits" />
<br />
<asp:Label cssClass="content8" runat="server" id="lblRating2" /> <img src="images/<%=strRatingimg%>.gif" style="vertical-align: middle;" alt="rating: <%=strRatingimg%>"> <asp:Label cssClass="content8" runat="server" id="lblranrating" />
 </div>
</div>
<!--End Random Recipe-->
<!--Begin Search Box-->
<div class="div8">
   <div class="div9"><asp:Label cssClass="content3" runat="server" id="lblheadersearch" /></div>
    <div class="div6">
     <asp:Label cssClass="content2" runat="server" id="lblsearch" />
     <form method="GET" action="search.aspx" style="margin-top: 0; margin-bottom: 0;">
     <input type="text" ID="find" Name="find" class="textbox" size="20" value="">
   <input type="submit" class="submit" ID="submit" name="submit" value="Search">
  </form>
 </div>
</div>
<!--End Search Box-->
<!--Begin 15 Most Popular Box-->
   <div class="div8">
   <div class="div9"><asp:Label cssClass="content3" runat="server" id="lblheadermostpopular" /></div>
   <div class="div6">
   <asp:DataList cssClass="hlink" id="RecipeTopMost" RepeatColumns="1" runat="server">
   <ItemTemplate>
  <span class="bluearrow2">�</span>&nbsp;
  <a class="dt" title="Category (<%# DataBinder.Eval(Container.DataItem, "Category") %>) - <%# DataBinder.Eval(Container.DataItem, "Name") %> recipe" href='<%# DataBinder.Eval(Container.DataItem, "ID", "recipedetail.aspx?id={0}") %>'>
  <%# DataBinder.Eval(Container.DataItem, "Name") %></a>
   </ItemTemplate>
  </asp:DataList>
 </div>
</div>
<!--End 15 Most Popular Box-->
<!--Begin 15 Newest Recipes-->
<div class="div8">
    <div class="div9"><asp:Label cssClass="content3" runat="server" id="lblheadernewest" /></div>
     <div class="div6">
      <asp:DataList cssClass="hlink" id="RecipeNew" RepeatColumns="1" runat="server">
       <ItemTemplate>
        <span class="bluearrow2">�</span>&nbsp;
         <a class="dt" title="Category (<%# DataBinder.Eval(Container.DataItem, "Category") %>) - <%#                   DataBinder.Eval(Container.DataItem, "Name") %> recipe" href='<%# DataBinder.Eval(Container.DataItem, "ID",                    "recipedetail.aspx?id={0}") %>'><%# DataBinder.Eval(Container.DataItem, "Name") %></a>
     </ItemTemplate>
   </asp:DataList>
 </div>
</div>
<!--End 15 Newest Recipes-->
 </td>
  </tr>
</table>
<div style="height: 45px;"></div>
<!--#include file="inc_footer.aspx"-->