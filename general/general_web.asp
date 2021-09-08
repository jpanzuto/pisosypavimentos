<!--#include file="conexion.asp"-->

<%
ip = Request.Servervariables("REMOTE_ADDR")
fecha = Now()
url = Request.ServerVariables("SERVER_NAME") + Request.ServerVariables("url") 
url_anterior = Request.Servervariables("HTTP_REFERER")
navegador = Request.Servervariables("HTTP_USER_AGENT")

Establecer_Conexion "", ""
Consulta ("insert into Estadisticas (ip, fecha, url, urlanterior, navegador) values ("&Quote(HTMLSafe(SQLSafe(ip)))&", "&Quote(HTMLSafe(SQLSafe(fecha)))&", "&Quote(HTMLSafe(SQLSafe(url)))&", "&Quote(HTMLSafe(SQLSafe(url_anterior)))&","&Quote(HTMLSafe(SQLSafe(navegador)))&")")
Cerrar_Conexion 
%>

    <link rel="stylesheet" href="general/estilo.css"  type="text/css">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <meta http-equiv='Content-Type' content='text/html; charset=UTF-8'>
    <meta charset="iso-8859-1">
   	<meta http-equiv='content-language' content='Spanish'>
	<meta http-equiv="X-UA-Compatible" content="IE=edge">
	<meta name="author" content="Juan Panzuto">
	<meta name="robots" content="all">    
    <meta name="verify-v1" content="W3GTp81TrMt2uFdAB4COhRzPdagReIA0ptLA4F1Gnz4="/>
	<meta name="msvalidate.01" content="3E1890B3DAFF9F0DE30E8205F38F6389"/>
    
    <script type="text/javascript">
	  var _gaq = _gaq || [];
	  _gaq.push(['_setAccount', 'UA-31077276-1']);
	  _gaq.push(['_trackPageview']);
	
	(function() {
		var ga = document.createElement('script'); ga.type = 'text/javascript'; ga.async = true;
		ga.src = ('https:' == document.location.protocol ? 'https://ssl' : 'http://www') + '.google-analytics.com/ga.js';
		
	var s = document.getElementsByTagName('script')[0]; s.parentNode.insertBefore(ga, s);
	  })();
	</script> 