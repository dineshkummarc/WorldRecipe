<script runat="server">
    
    'Get the database server Map Path - This is where you change the database name and path
    Function DB_Path()
        
       if instr(Context.Request.ServerVariables("PATH_TRANSLATED"),"Recipes") then
            DB_Path = System.Web.HttpContext.Current.Server.MapPath("App_Data/recipedb.mdb;")
            'DB_Path = System.Web.HttpContext.Current.Server.MapPath("App_Data|\recipedb.mdb;")

        Else
            DB_Path = System.Web.HttpContext.Current.Server.MapPath("App_Data/recipedb.mdb;")
            'DB_Path = System.Web.HttpContext.Current.Server.MapPath("App_Data|\recipedb.mdb;")

        End If

    End Function

</script>
