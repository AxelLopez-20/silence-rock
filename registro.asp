<html>
<body>
<%@ LANGUAGE="VBSCRIPT" %>
<%
option explicit
Dim conn,nom,email,coments,city,nac

nom=Request.Form("nombre")
email=request.Form("correo")
coments=Request.Form("coms")
city= Request.From("ciudad")
nac= Request.From("nacimiento")

set conn=Server.CreateObject("ADODB.connection")
Conn.open "provider=Microsoft.Jet.OLEDB.4.0;Data Source=registro.mdb"
Conn.execute "INSERT INTO datoscliente(nombre,correo,ciudad,nacimiento) values('"& nom & "','"& email & "','"& city & "','"& nac &"')"
Conn.close
set conn=nothing
responde.redirect("gracias.html")

%>
</body>
</html>