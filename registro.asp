<html>
<body>
<%@ LANGUAGE="VBSCRIPT" %>
<%
option explicit
Dim conn,nom,email,city,nac

nom=Request.Form("nombre")
email=request.Form("correo")
city= Request.From("ciudad")
nac= Request.From("nacimiento")

set conn=Server.CreateObject("ADODB.connection")
Set Tabla = Server.CreateObject("ADODB.Recordset")
Conn.Open "Driver={Microsoft Access Driver (*.mdb)}; " & "Dbq=" & Server.MapPath("registro.mdb")
Conn.execute "INSERT INTO datoscliente(nombre,correo,ciudad,nacimiento) values('"& nom & "','"& email & "','"& city & "','"& nac &"')"
Conn.close
set conn=nothing
responde.redirect("gracias.html")

%>
</body>
</html>