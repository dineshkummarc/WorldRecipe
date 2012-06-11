<%@ Page Language="VB" Debug="true" %>
<%@ import Namespace="System.Data" %>
<%@ import Namespace="System.Data.Oledb" %>

<script runat="server">

  'Handle page load event
  Sub Page_Load(Sender As Object, E As EventArgs)            
        
         'Call count total number of recipes
         CountNumberRecipes()

         'Call NewestRecipes Sub - Display 15 newest recipes
         NewestRecipes()

         'Call MostPopular Sub - Display 15 most popular recipes
         MostPopular() 

         'Call display sort category links
         SortCategoryLink()
        
         'Call risplay recipe name A to Z
         DisplayLetter()

         DisplayCatName()

         GetCatCategory()
       
         'Call Random Recipe
         RandomRecipeNumber()
         RandomRecipe()

         lblletter.Text = "Recipe Name List A-Z:"
         lblsearch.Text = "Search recipe by name,author or country of origin i.e.(Filipino,chinese)"
         lblheadersearch.text = "Search Recipe"
         lblheadermostpopular.text = "15 Most Popular"
         lblheadernewest.text = "15 New Recipes"
         lblsortcat.text = "Sort Category"  
         lblheaderrandom.text = "Featured Recipe" 
         lblCatNameHeader.text = "Choose Category"      

         HyperLink1.NavigateUrl = "index.aspx"
         HyperLink1.Text = "Recipe Home"
         HyperLink1.ToolTip = "Back to recipe homepage"


        If Not Page.IsPostBack Then
                  
           ViewState("Start") = 0
          
           'Call the bindlist - show data
           BindList()
        
       End If 
     
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
          
        Dim strCatTop as integer = Request.QueryString("catid")

strSQL = "SELECT Top 15 ID,Name,HITS,Category FROM Recipes WHERE CAT_ID = " & Replace(strCatTop, "'", "''") & " AND LINK_APPROVED = 1 Order By HITS DESC"
  
         'Call Open database - connect to the database      
         DBconnect()

         Dim RecipeAdapter as New OledbDataAdapter(objCommand)
         Dim dts as New DataSet()
         RecipeAdapter.Fill(dts, "ID")

         RecipeTop.DataSource = dts.Tables("ID").DefaultView
         RecipeTop.DataBind()

 End Sub



  'Bindlist Datasource
  Sub BindList()


         Dim intpageSize As Integer 
         Dim CategoryID as Integer
         Dim recipesqlorderby as string
         Dim action as string 
         Dim orderby as string
         Dim sortname as string
         Dim sqlorderby as string
         Dim strCaption as string
         Dim RcdCount As Integer
        

         CategoryID = Request.QueryString("catid")


         action = Request.QueryString("action")
         action = "Date"
         orderby = "DESC"


    If(CStr(Request.QueryString("sid")) = "0") Then
	    action = "Date"
            strCaption = ""

        ElseIf(CStr(Request.QueryString("sid")) = "") Then
	    action = "Date"
            strCaption = "" 

        ElseIf(CStr(Request.QueryString("sid")) < "0") Then
	    action = "Date"
            strCaption = ""

        ElseIf(CStr(Request.QueryString("sid")) = "1") Then
            action = "NO_RATES"
            strCaption = "Sorted by: Highest Rated"

        ElseIf(CStr(Request.QueryString("sid")) = "2") Then
            action = "HITS"
            strCaption = "Sorted by: Most Popular"

        ElseIf(CStr(Request.QueryString("sid")) = "3") Then
            action = "Date"
            strCaption = "Sorted by: Newest"

        ElseIf(CStr(Request.QueryString("sid")) = "4") Then
	    action = "Name"
            strCaption = "Sorted by: Name ASC"
        
        ElseIf(CStr(Request.QueryString("sid")) > "4") Then
	    action = "Date"
            strCaption = ""             

   End If
   
  'If order ASC or DESC equals blank then grab the default value 
  'default value is set to Date field, else append from querystring OB = 1 ASC or 2 Desc
  If(Cstr(Request.QueryString("ob")) <> "") Then
	orderby = Cstr(Request.QueryString("ob"))
  End If

   'Sort by whether Ascending or Descending
   If orderby = "1" Then
	orderby = "ASC"

    ElseIf orderby = "2" Then
	orderby = "DESC"

    ElseIf orderby > "2" Then
	orderby = "DESC"

    ElseIf orderby = "0" Then
	orderby = "DESC"

    ElseIf orderby < "0" Then
	orderby = "DESC"

  End If

         sqlorderby = " " & action & " " & orderby
         recipesqlorderby = "Date"
         if (sqlorderby <> "") then recipesqlorderby = sqlorderby
         
         'SQL statement display category
strSQL = "SELECT *, (RATING/NO_RATES) AS Rates FROM Recipes WHERE CAT_ID = " & Replace(CategoryID, "'", "''") & " AND LINK_APPROVED = 1 ORDER BY " & Replace(recipesqlorderby, "'", "''") & "" 

         'Call Open database - connect to the database      
         DBconnect()
       
         Dim RecipeAdapter as New OledbDataAdapter(objCommand)
         Dim dts as New DataSet()

         intStart = ViewState("Start")
         ViewState("pageSize") = 10

         RecipeAdapter.Fill(dts, intStart, ViewState("pageSize"), "ID") 

         RecipeCat.DataSource = dts         
         RecipeCat.DataBind() 

         lblcaption.text = strCaption     

 End Sub



'Display sort category links
 Sub SortCategoryLink()

        Dim strSid as integer
        Dim strCatid as integer

        strCatid = Request.QueryString("catid")
        strSid = Request.QueryString("sid")

        If strSid = "2" Then
           LinkMostPopular.Enabled = False
           LinkMostPopular.Text = "Most Popular"
        Else 
           LinkMostPopular.NavigateUrl = "category.aspx?catid=" & strCatid & "&sid=" & 2
           LinkMostPopular.Text = "Most Popular"
           LinkMostPopular.ToolTip = "Sort Category by Most Popular Recipes"
        End if

        If strSid = "1" Then
           LinkHighestRated.Enabled = False
           LinkHighestRated.Text = "Highest Rated"
        Else 
           LinkHighestRated.NavigateUrl = "category.aspx?catid=" & strCatid & "&sid=" & 1
           LinkHighestRated.Text = "Highest Rated"
           LinkHighestRated.ToolTip = "Sort Category by Newest Rated Recipes"
        End if

        If strSid = "3" Then
           LinkNewest.Enabled = False
           LinkNewest.Text = "Newest"
        Else 
           LinkNewest.NavigateUrl = "category.aspx?catid=" & strCatid & "&sid=" & 3
           LinkNewest.Text = "Newest"
           LinkNewest.ToolTip = "Sort Category by Newest Recipes"
        End if

        If strSid = "4" Then
           LinkName.Enabled = False
           LinkName.Text = "Name ASC"
        Else 
           LinkName.NavigateUrl = "category.aspx?catid=" & strCatid & "&sid=" & 4 & "&ob=" & 1
           LinkName.Text = "Name ASC"
           LinkName.ToolTip = "Sort Category by Recipe Name ASC"
        End if

 End Sub



  'Page level error handling - If the page encounter an error, redirect to the custom error page
  Protected Overrides Sub OnError(ByVal e As System.EventArgs)

    Server.Transfer("error.aspx")

  End Sub



 'Count the number of recipes in the selected category
 Sub CountNumberRecipes()

        Dim getcatid as string
        getcatid = Request.QueryString("catid")

  strSQL = "SELECT Count(CAT_ID) FROM Recipes WHERE CAT_ID = " & Replace(getcatid, "'", "''")

       'Call Open database - connect to the database      
        DBconnect()

        objCommand.Connection.Open()
lblrcdcount.Text = "There are&nbsp;" & "( " & objCommand.ExecuteScalar() & " )" & "&nbsp;recipes" 
        objCommand.Connection.Close()

 End Sub

  

  'Set the paging next link
  Sub Next_Click(Sender As Object, E As EventArgs)

       
       Dim dlistcount As Integer = RecipeCat.Items.Count

       intStart = ViewState("Start") + ViewState("pageSize")
       ViewState("Start") = intStart
     
   If dlistcount < ViewState("pageSize") Then
          ViewState("Start") = ViewState("Start") - ViewState("pageSize")
   End If

     'Call Bindlist sub and display data
     BindList()

 End Sub



 'Set the paging previous link
 Sub Prev_Click(Sender As Object, E As EventArgs)

      intStart = ViewState("Start") - ViewState("pageSize")
      ViewState("Start") = intStart
    
        If intStart <= 0 Then 
          ViewState("Start") = 0
        End If

     'Call Bindlist sub and display data
     BindList() 

 End Sub 



 'Jump back to the first page
 Sub FirstPage(Sender As Object, E As EventArgs)

        If intStart <= 0 Then 
          ViewState("Start") = 0
        End If

     'Call Bindlist - show data
     BindList() 

 End Sub 



  'Pulls a Random number for selecting a random recipe
  Sub RandomRecipeNumber()

        'It connects to database
        strSQL = "SELECT CAT_ID FROM Recipes WHERE CAT_ID = " & Request.QueryString("catid") 

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

strSQL = "SELECT ID,CAT_ID,Category,Name,Author,Date,HITS,RATING,NO_RATES, (RATING/NO_RATES) AS Rates FROM Recipes WHERE CAT_ID = " & Request.QueryString("catid") 

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
        lblrancat2.text = objDataReader("Category")
        strRanRating = FormatNumber(objDataReader("Rates"), 1,  -2, -2, -2)
        lblranrating.Text = "(" & strRanRating & ")"
        strRatingimg = FormatNumber(objDataReader("Rates"), 1,  -2, -2, -2)

        LinkRanName.NavigateUrl = "recipedetail.aspx?id=" & objDataReader("ID")
        LinkRanName.Text = objDataReader("Name")
        LinkRanName.ToolTip = "View" & " - " & objDataReader("Name") & " - " & "recipe"

        objDataReader.Close()
        objConnection.Close()

  End Sub

 
  
  'Display category name
  Sub DisplayCatName()

 strSQL = "SELECT Category FROM Recipes WHERE CAT_ID =" & Request.QueryString("catid")
    
            Dim objDataReader as OledbDataReader
            
            'Call Open database - connect to the database
            DBconnect()

            objConnection.Open()
            objDataReader  = objCommand.ExecuteReader()
    
            'Read data
            objDataReader.Read()

          lblCategoryName.text = objDataReader("Category")

          'Close database connection for the objDataReader
          objDataReader.Close()
          objConnection.Close() 


  End Sub



 'Database connection string
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
    Public intStart As Integer
    Public strRatingimg as Integer
    Public iRandomRecipe as integer

</script>

<!--#include file="inc_databasepath.aspx"-->

<!--#include file="inc_header.aspx"-->

<table border="0" cellpadding="0" cellspacing="0" width="100%">
  <tr>
    <td width="16%" rowspan="5" valign="top">
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
    <td width="68%">
<div style="margin-left: 12px;">
<asp:HyperLink tooltip="Back to recipe homepage" id="HyperLink1" cssClass="dtcat" runat="server" />&nbsp;<span class="bluearrow">�</span>&nbsp;<asp:Label cssClass="content10" runat="server" id="lblCategoryName" />
</div>
</td>
    <td width="16%" rowspan="5" valign="top">
<!--Begin Random Recipe-->
    <div class="div8">
<div class="div9"><asp:Label cssClass="content3" runat="server" id="lblheaderrandom" /></div>
 <div class="div6">
<span class="bluearrow2">�</span>
<asp:HyperLink id="LinkRanName" cssClass="dtcat" runat="server" />
<br />
<asp:Label cssClass="content8" runat="server" id="lblRancategory" /> <asp:Label cssClass="content8" runat="server" id="lblrancat2" />
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
<!--15 Most Popular Recipe-->
<div class="div8">
 <div class="div9"><asp:Label cssClass="content3" runat="server" id="lblheadermostpopular" /></div>
  <div class="div6">
   <asp:DataList cssClass="hlink" id="RecipeTop" RepeatColumns="1" runat="server">
   <ItemTemplate>
<span class="bluearrow2">�</span>&nbsp;
<a class="dt" title="Category (<%# DataBinder.Eval(Container.DataItem, "Category") %>) - <%# DataBinder.Eval(Container.DataItem, "Name") %> recipe" href='<%# DataBinder.Eval(Container.DataItem, "ID", "recipedetail.aspx?id={0}") %>'>
<%# DataBinder.Eval(Container.DataItem, "Name") %></a>
    </ItemTemplate>
  </asp:DataList>
 </div>
</div>
<!--End 15 Most Popular Recipe-->
<!--Begin 15 Newest Recipes-->
  <div class="div8">
  <div class="div9"><asp:Label cssClass="content3" runat="server" id="lblheadernewest" /></div>
 <div class="div6">
<asp:DataList cssClass="hlink" id="RecipeNew" RepeatColumns="1" runat="server">
   <ItemTemplate>
<span class="bluearrow2">�</span>&nbsp;
<a class="dt" title="Category (<%# DataBinder.Eval(Container.DataItem, "Category") %>) - <%# DataBinder.Eval(Container.DataItem, "Name") %> recipe" href='<%# DataBinder.Eval(Container.DataItem, "ID", "recipedetail.aspx?id={0}") %>'>
<%# DataBinder.Eval(Container.DataItem, "Name") %></a>
   </ItemTemplate>
  </asp:DataList>
 </div>
</div>
<!--End 15 Newest Recipes-->
</td>
  </tr>
  <tr>
    <td width="68%">
<div style="margin-left: 16px;">
<asp:Label cssClass="content2" ID="lblrcdcount" runat="server" />&nbsp;&nbsp;<asp:Label cssClass="content2" ID="lblcaption" runat="server" />
</div> 
</td>
  </tr>
<tr>
    <td width="68%">
<div style="padding: 2px; text-align: center; border: 1px dashed #e5e5e5; margin-bottom: 10px; margin-top: 10px; margin-left: 26px; margin-right: 26px;">
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
    <td width="68%" align="right">
<div class="divsort">
<asp:Label class="sortcat" runat="server" id="lblsortcat" />:
<asp:HyperLink id="LinkMostPopular" cssClass="dt" runat="server" /> |
<asp:HyperLink id="LinkHighestRated" cssClass="dt" runat="server" /> |
<asp:HyperLink id="LinkNewest" cssClass="dt" runat="server" /> |
<asp:HyperLink id="LinkName" cssClass="dt" runat="server" />
</div>
</td>
  </tr>
  <tr>
    <td width="68%" valign="top">
<asp:DataList width="100%" id="RecipeCat" RepeatColumns="1" runat="server">
      <ItemTemplate>
    <div class="divwrap">
       <div class="divhd">
<span class="bluearrow">�</span>&nbsp;
<a class="dtcat" title="View <%# DataBinder.Eval(Container.DataItem, "Name") %> recipe" href='<%# DataBinder.Eval(Container.DataItem, "ID", "recipedetail.aspx?id={0}") %>'><%# DataBinder.Eval(Container.DataItem, "Name") %></a>
</div> 
<div class="divbd">
Category:&nbsp;<%# DataBinder.Eval(Container.DataItem, "Category") %>
<br />
Submitted by:&nbsp;<%# DataBinder.Eval(Container.DataItem, "Author") %>
<br />
Rating:&nbsp;<img src="images/<%# FormatNumber((DataBinder.Eval(Container.DataItem, "Rates")), 1,  -2, -2, -2) %>.gif" style="vertical-align: middle;" alt="Rating <%# FormatNumber((DataBinder.Eval(Container.DataItem, "Rates")), 1, -2, -2, -2) %>">&nbsp;( <%# FormatNumber((DataBinder.Eval(Container.DataItem, "Rates")), 1, -2, -2, -2) %> ) by <%# DataBinder.Eval(Container.DataItem, "NO_RATES") %> users
<br />
Added: <%# FormatDateTime(DataBinder.Eval(Container.DataItem, "Date"),vbShortDate) %>
<br />
Hits: <%# DataBinder.Eval(Container.DataItem, "HITS") %>
       </div>
</div>
      </ItemTemplate>
  </asp:DataList>
<form runat="server" style="margin-top: 0px; margin-bottom: 0px;">
<div style="padding: 5px; margin-bottom: 35px;" class="content2">
<asp:LinkButton cssClass="dt" id="first10" OnClick="FirstPage" runat="server"><< First Page</asp:LinkButton> |
<asp:LinkButton cssClass="dt" id="ClickPrev" OnClick="Prev_Click" runat="server"><< Previous Page</asp:LinkButton> |
<asp:LinkButton cssClass="dt" id="ClickNext" OnClick="Next_Click" runat="server">Next Page >></asp:LinkButton> 
</div>
</form>
</td>
  </tr>
</table>
<!--#include file="inc_footer.aspx"-->