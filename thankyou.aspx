<%@ Page Language="VB" Debug="true" %>
<%@ import Namespace="System.Data" %>
<%@ import Namespace="System.Data.Oledb" %>

<script runat="server">

  'Handle the page load event
  Sub Page_Load(Sender As Object, E As EventArgs)

         'Call TotalRecipeCount Sub routine - get the number recipes
         TotalRecipeCount()

         'Call CategoryCount Sub routine - get the number of categories
         CategoryCount()

         'Call NewestRecipes Sub - Display 15 newest recipes
         NewestRecipes()

         'Call MostPopular Sub - Display 15 most popular recipes
         MostPopular()

         'Call Random Recipe
         RandomRecipeNumber()
         RandomRecipe()

         lblletter.Text = "Thank you for submitting a recipe to our website"
         lblsearch.Text = "Search recipe by name,author or country of origin i.e.(Filipino,chinese)"
         lblheadermostpopular.text = "15 Most Popular"
         lblheadernewest.text = "15 New Recipes"
         lblheadersponsor.text = "Sponsors"
         lblheaderrandom.text = "Featured Recipe"    

 End Sub


 'Sub Display 15 newest recipes
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



  'Get the total number of recipes
  Sub TotalRecipeCount()

        Dim CmdCount As New OleDbCommand("Select Count(ID) From Recipes", New OleDbConnection(strConnection))
        CmdCount.Connection.Open()
        lbltotalRecipe.Text = "There are &nbsp;" & CmdCount.ExecuteScalar() & "&nbsp;recipes in&nbsp;"
        CmdCount.Connection.Close()

  End Sub




 'Get the total number of categories
 Sub CategoryCount()

        Dim CmdCount2 As New OleDbCommand("Select Count(CAT_ID) From RECIPE_CAT", New OleDbConnection(strConnection))
        CmdCount2.Connection.Open()
        lbltotalCat.Text = CmdCount2.ExecuteScalar() & "&nbsp;categories"
        CmdCount2.Connection.Close()

 End Sub




 'Display 15 most popular recipes
 Sub MostPopular()
         
strSQL = "SELECT Top 15 ID,Name,HITS,Category FROM Recipes Where LINK_APPROVED = 1 Order By HITS DESC"

         'Call Open database - connect to the database
         DBconnect()

         Dim RecipeAdapter as New OledbDataAdapter(objCommand)
         Dim dts as New DataSet()
         RecipeAdapter.Fill(dts, "ID")

         RecipeTop.DataSource = dts.Tables("ID").DefaultView
         RecipeTop.DataBind()

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


    'Declare public so it will accessible in all subs
    Public strDBLocation = DB_Path()
    Public strConnection as string = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strDBLocation
    Public objConnection
    Public objCommand
    Public strSQL as string
    Public strRatingimg as Integer
    Public iRandomRecipe as integer
    

</script>

<!--#include file="inc_databasepath.aspx"-->
       
<!--#include file="inc_header.aspx"-->

<table border="0" cellpadding="0" cellspacing="0" width="100%">
 <tr>
    <td width="15%" valign="top" align="left">
    <!--#include file="inc_navmenu.aspx"-->
<span class="content3">br</span>
<!--Begin 15 Newest Recipes-->
  <div class="div8">
<div class="div9"><asp:Label cssClass="content3" runat="server" id="lblheadernewest" /></div>
 <div class="div6">
<asp:DataList cssClass="hlink" id="RecipeNew" RepeatColumns="1" runat="server">
   <ItemTemplate>
<span class="bluearrow2">»</span>&nbsp;
<a class="dt" title="Category (<%# DataBinder.Eval(Container.DataItem, "Category") %>) - <%# DataBinder.Eval(Container.DataItem, "Name") %> recipe" href='<%# DataBinder.Eval(Container.DataItem, "ID", "recipedetail.aspx?id={0}") %>'>
<%# DataBinder.Eval(Container.DataItem, "Name") %></a>
      </ItemTemplate>
  </asp:DataList>
 </div>
</div>
<!--End 15 Newest Recipes-->
    </td>
    <td width="70%" valign="top">
<div style="padding-left: 60px; padding-top: 16px; margin-bottom: 6px;">
<asp:Label cssClass="content2" runat="server" id="lblsearch" />
<form method="GET" action="search.aspx" style="margin-top: 0; margin-bottom: 0;">
 <input type="text" ID="find" Name="find" class="textbox" size="20" value="">
 <input type="submit" class="submit" ID="submit" name="submit" value="Search">
 </form>
</div>
<div style="padding-left: 60px; padding-top: 12px;"><asp:Label cssClass="content2" runat="server" id="lbltotalRecipe" /><asp:Label cssClass="content2" runat="server" id="lbltotalCat" /></div>
<br />
<br />
<br />
<div style="padding: 2px; text-align: center; margin-left: 26px; margin-right: 26px;">
<asp:Label cssClass="corange" runat="server" id="lblletter" />
<br />
<span class="content2">Recipe has been successfully added to our database.</span>
<br />
<span class="content2">An email was sent to the Webmaster for recipe review and approval. Your recipe will be posted with in 12 hrs upon approval.</span>
</div>
</div>
<!--Begin Thank you-->
<table border="0" cellpadding="2" align="center" cellspacing="2" width="60%">
  <tr>
    <td width="100%" align="center" valign="top">
<div style="margin-top:25px;">
<a class="dt" href="submitrecipe.aspx">Submit another recipe</a>
<br />
<a class="dt" href="index.aspx">Back to recipe homepage</a>
</div>
</td>
  </tr>
</table>
<!--End Thank you-->
    </td>
    <td width="15%" valign="top" valign="top" align="left">
<!--Begin Random Recipe-->
    <div class="div8">
<div class="div9"><asp:Label cssClass="content3" runat="server" id="lblheaderrandom" /></div>
 <div class="div6">
<span class="bluearrow2">»</span>
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
<!--Begin 15 Most Popular-->
    <div class="div8">
<div class="div9"><asp:Label cssClass="content3" runat="server" id="lblheadermostpopular" /></div>
 <div class="div6">
<asp:DataList cssClass="hlink" id="RecipeTop" RepeatColumns="1" runat="server">
   <ItemTemplate>
<span class="bluearrow2">»</span>&nbsp;
<a class="dt" title="Category (<%# DataBinder.Eval(Container.DataItem, "Category") %>) - <%# DataBinder.Eval(Container.DataItem, "Name") %> recipe" href='<%# DataBinder.Eval(Container.DataItem, "ID", "recipedetail.aspx?id={0}") %>'>
<%# DataBinder.Eval(Container.DataItem, "Name") %></a>
      </ItemTemplate>
  </asp:DataList>
 </div>
</div>
<!--End 15 Most Popular-->
<!--Begin Sponsors Box-->
    <div class="div8">
<div class="div9"><asp:Label cssClass="content3" runat="server" id="lblheadersponsor" /></div>
 <div class="div6">
<a title="Visit ex-designz.net" href="http://www.ex-designz.net"><Img border="0" src="http://www.ex-designz.net/ex-designs_sm.gif" alt="Visit ex-designz.net" width="88" height="31"></a>
<br />
<a title="Isnare Article directory" href="http://www.isnare.com/" target="_blank"><Img border="0" src="http://www.isnare.com/banners/120x60-animated.gif" alt="Isnare Article directory" width="120" height="60"></a>
 </div>
</div>
<!--End Sponsor Box-->
</td>
  </tr>
</table>
<div style="margin-top: 80px;"></div>
<!--#include file="inc_footer.aspx"-->