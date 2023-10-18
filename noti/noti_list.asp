<%@Language=VBScript CodePage=65001%>
<!--#include virtual=/sb/inc/vbutil.inc-->
<%
  Response.CodePage = 65001
  Response.CharSet = "utf-8"

%>
<html>
	<head>
			<meta http-equiv="content-type" content="text/html;charset=utf-8" />
			<link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/4.6.3/css/font-awesome.min.css">
			<link rel="stylesheet" href="http://mdev.gvg.co.kr/css/master.css">
			<link rel="stylesheet" href="https://ajax.googleapis.com/ajax/libs/jqueryui/1.12.1/themes/smoothness/jquery-ui.css">
			<link rel="stylesheet" type="text/css" href="http://mdev.gvg.co.kr/css/datatables.min.css"/>
 			<link href="https://fonts.googleapis.com/icon?family=Material+Icons" rel="stylesheet">
 			<style>
	 			.table_datagrid td{text-align:left;}
	 		</style>
			<script src="https://ajax.googleapis.com/ajax/libs/jquery/3.2.1/jquery.min.js"></script>
			<script src="https://ajax.googleapis.com/ajax/libs/jqueryui/1.12.1/jquery-ui.min.js"></script>
			<script type="text/javascript" src="http://mdev.gvg.co.kr/jquery/datatables.min.js"></script>
			<script type="text/javascript" src="http://mdev.gvg.co.kr/jquery/fixeddatatable.js"></script>
			<script type="text/javascript" src="http://mdev.gvg.co.kr/jquery/jquery.form.js"></script>
			<script src="http://mdev.gvg.co.kr/jquery/onload.js"></script>
			<script src="http://mdev.gvg.co.kr/jquery/master.js"></script>
			<script>
				function addData() {
					window.location.href ='set_noti_info.html';
				}
				
				function fnSubmit(noti_id) {
					var r = confirm("삭제?");
					if (r == true) {
							$.get('del_noti_info.html', { noti_id: noti_id })
								.done(function(response) {
									location.reload();
								})
								.fail(function(response) {
									alert("삭제 실패");
								});
						}
				}
			</script>
		</head>
		<body>
			<ul class="manual">
				<li>tb_noti_master 에 기본정보 입력화면 구성</li>
				<li>*_flag 필드는 1/0 1:발송 0:제외</li>
				<li>due_date: 발송일 지정, due 발송차수 지정 1~6 (차수별 발송시간은 차후설정)</li>
			</ul>
			<h1 style="display: inline-block;">NOTIFICATION LIST </h1>
			<a class="btn5 biggest" onclick=addData();>ADD</a>
			
            <%
			 sql = "SELECT * FROM tb_noti_master ORDER BY noti_id DESC" 
			set rs=getRs(sql)
			'Response.Write("<p>쿼리: " & sql & "</p>")

			Response.Write "<br><table class='table_datagrid'>"
				'제목
				While Not rs.EOF 
					Response.Write "<tr>"
					Response.Write "<th>" & rs.Fields("noti_id").Value & "</th>"
					Response.Write "<td><a href='get_noti_form.html?noti_id=" & rs.Fields("noti_id").Value & "'>" & rs.Fields("title").Value & "</a> <a class='btn3 biggest' onclick='fnSubmit(" & rs.Fields("noti_id").Value & ");'>DEL</a></td>"
					Response.Write "</tr>"
					rs.MoveNext      
				Wend
		
			Response.Write "</table>"

			action = Request.Form("action")			
			noti_id = Request.Form("noti_id")
			title = Request.Form("title")
			noti_id = Replace(noti_id, ",", "")
			
			' noti_id = Request.QueryString("noti_id")
 			' title = Request.QueryString("title")
			' Response.Write ("<h1>noti_id / QueryString "& noti_id &"</h1>")
			' Response.Write ("<h1>QueryString / QueryString "& title &"</h1>")
				Set conn = Server.CreateObject("ADODB.Connection")
				conn.ConnectionString = "Provider=SQLOLEDB;Data Source=sql16ssd-004.localnet.kr,1433;Initial Catalog=gvgurbanstoreb_planb;User ID=gvgurbanstoreb_planb;Password=rnjswodls1030;"
				conn.Open

			If action = "set" Then 
				Response.Write ("<h1>"& action &"</h1>")
			 	Response.Write("<p>set action</p>")
				sql1= "INSERT INTO tb_noti_master (title) VALUES ('" & title & "')"
				Response.Write ("<h2> SQL: "& sql1 &"</h2>")
				Response.Write ("<h2> title: "& title &"</h2>")
				Set cmd_insert= Server.CreateObject("ADODB.Command")
				cmd_insert.ActiveConnection= conn
				cmd_insert.CommandText= sql1
				cmd_insert.Execute 
				Response.Redirect("noti_list.asp")

			ElseIf action = "del" Then
				Response.Write("<p>del action</p>")
				sql = "DELETE FROM tb_noti_master WHERE noti_id = " & noti_id  
				Set conn = Server.CreateObject("ADODB.Connection")
				conn.ConnectionString = "Provider=SQLOLEDB;Data Source=sql16ssd-004.localnet.kr,1433;Initial Catalog=gvgurbanstoreb_planb;User ID=gvgurbanstoreb_planb;Password=rnjswodls1030;"
				conn.Open
				Set cmd = Server.CreateObject("ADODB.Command")
				cmd.ActiveConnection = conn
				cmd.CommandText = sql
				cmd.Execute

			ElseIf action = "update, update" Then
			 	Response.Write("<p>update action</p>")
				sql1 = "UPDATE tb_noti_master SET title = '" & title & "' WHERE noti_id = " & noti_id 
				Response.Write("<p>sql1: " & sql1 & "</p>")
				Set conn = Server.CreateObject("ADODB.Connection")
				conn.ConnectionString = "Provider=SQLOLEDB;Data Source=sql16ssd-004.localnet.kr,1433;Initial Catalog=gvgurbanstoreb_planb;User ID=gvgurbanstoreb_planb;Password=rnjswodls1030;"
				conn.Open
				Set cmd = Server.CreateObject("ADODB.Command")
				cmd.ActiveConnection = conn
				cmd.CommandText = sql1
				cmd.Execute
				Response.Redirect("noti_list.asp")

			End If
			
			conn.Close
			Set conn = Nothing 
			'Response.Redirect "noti_list.html"

			Response.Write("<p>noti_id???:" & noti_id & "</p>")
			Response.Write("<p>title: " & title & "</p>")
			Response.Write ("<h1>"& action &"</h1>")
			
			%>
		</body>
</html>