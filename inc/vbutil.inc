<%
Function getRs(sql)
      Dim conn, rs
    Set conn = Server.CreateObject("ADODB.Connection")
    
    ' 데이터베이스 연결 설정
    conn.ConnectionString = "Provider=SQLOLEDB;Data Source=sql16ssd-004.localnet.kr,1433;Initial Catalog=gvgurbanstoreb_planb;User ID=gvgurbanstoreb_planb;Password=rnjswodls1030;"
    conn.Open
    Set rs = Server.CreateObject("ADODB.Recordset")
    rs.Open sql, conn

    Set getRs = rs
End Function
%>
