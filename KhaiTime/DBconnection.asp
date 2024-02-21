<%
    ConnectionString = "Provider=SQLOLEDB;Data Source=khaiphan.database.windows.net;Initial Catalog=KhaiTime;User Id=khaiphan; Password = Phankhai11423!;"
    set conn = Server.CreateObject("ADODB.Connection")
    conn.Open ConnectionString
%>  