<html>
<body>
<%@ LANGUAGE="VBSCRIPT" %>
<%
option explicit
Dim conn,nom,email,coments

nom=Request.Form("nombre")
email=request.Form("correo")
coments=Request.Form("coms")

set conn=Server.CreateObject("ADODB.connection")
Conn.open "provider=Microsoft.Jet.OLEDB.4.0;Data Source=/registro"
Conn.execute "INSERT INTO datoscliente(nombre,correo,ciudad,nacimiento) values('"& nombre & "','"& correo & "','"& ciudad & "','"& nacimiento &"')"
Conn.close
set conn=nothing
responde.redirect("gracias.html")

%>
</body>
</html>