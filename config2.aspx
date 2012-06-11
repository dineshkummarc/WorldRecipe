<script runat="server">
    
    'Get the database server Map Path
    Function DB_Path()
        
       if instr(Context.Request.ServerVariables("PATH_TRANSLATED"),"Recipes") then
          DB_Path = System.Web.HttpContext.Current.Server.MapPath("/db/recipedb.mdb")
       else
          DB_Path = System.Web.HttpContext.Current.Server.MapPath("/db/recipedb.mdb")
       end if

    End Function

    Public strDBLocation = DB_Path()
    Public strConnection as string = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strDBLocation
    Public objConnection
    Public objCommand
    Public objDataReader as OledbDataReader


    'Database connection
    Sub DataBase_Connect(strSQL)
    
        objConnection = New OledbConnection(strConnection)
        objConnection.Open()
        objCommand = New OledbCommand(strSQL, objConnection)
        objDataReader  = objCommand.ExecuteReader()
    
    End Sub
    
    'Database disconnect
    Sub DataBase_Disconnect()
    
        objDataReader.Close()
        objConnection.Close()
    
    End Sub

</script>

