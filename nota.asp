<!DOCTYPE HTML> 
<html lang="es">
<!-- #include virtual="/general/general_web.asp" -->
<%
Dim mostrar 'cantidad de registros a mostrar por página
Dim cant_paginas 'cantidad de páginas que recibimos
Dim pagina_actual 'La página que mostramos
Dim registro_mostrado 'Contador utilizado para mostrar las páginas
Dim I 'Variable Loop

mostrar = 1

strsql = "SELECT * FROM Notas Order By id"

Set Conexion = Server.CreateObject("ADODB.Connection")
Conexion.Open "DRIVER={Microsoft Access Driver (*.mdb)}; DBQ="&Server.MapPath("/notas/bd_notas.mdb")

' Creamos el RecordSet y definimos la cantidad de registros a mostrar
Set RS = Server.CreateObject("ADODB.Recordset")
RS.PageSize = mostrar
RS.CacheSize = mostrar

RS.Open strSQL, Conexion,3,1

'contamos las páginas que se formaron con la variable mostrar.
cant_paginas = RS.PageCount

' IF para saber que página mostrar
If Request.QueryString("pagina") = "" Then
	pagina_actual = cant_paginas
	Else
	pagina_actual = CInt(Request.QueryString("pagina"))
End If

' Si el pedido de página cae afuera del rango,
' lo modificamos para que caiga adentro
' Estas lineas no deberian ejecutarse
If pagina_actual > cant_paginas Then pagina_actual = cant_paginas
If pagina_actual < 1 Then pagina_actual = 1

' Si la cantidad de páginas da 0 es que no hay registros... por eso este IF
If cant_paginas = 0 Then
	Response.Write "No hay registros para mostrar"
Else
	' Nos movemos a la página elegida
	RS.AbsolutePage = pagina_actual

%>

<head>   
    <title><% Response.Write RS.Fields(1) %></title>
	<meta name="title" content="<% Response.Write RS.Fields(1)%>"> 
    <meta name="keywords" content="<% Response.Write RS.Fields(1)%>">
	<meta name="description" content="<% Response.Write RS.Fields(1)%>">
    
    <link rel="stylesheet" href="/general/estilo.css"  type="text/css"><!--[if lt IE 9]>
    <script src="http://html5shiv.googlecode.com/svn/trunk/html5.js"></script>
    <![endif]-->
    <meta content="text/html; charset=iso-8859-1" http-equiv=Content-Type>
   	<meta http-equiv='content-language' content='Spanish'/>
	<meta name="author" content="J.I.Panzuto">
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

</head>

<body>
<div class="container">
  <header> <!--  Encabezado -->
	<!-- #include virtual="/general/flash_encabezado.asp" -->
  </header>
  
  <div class="sidebar1">
	<!-- #include virtual="/general/menu_navegacion.asp" -->
    
    <aside id="izquierda">
      <p>Construccion y refaccion de concreto, pulido de pisos.</p>
	  <p>&nbsp;</p>
      <p>&nbsp;</p>
</aside>
  </div>
  
  <section>
  <article class="content">
    <hgroup>
    <h1>Nota</h1>
    <h4><hr align="center"/><% Response.Write RS.Fields(1) %>
	<!-- #include virtual="/general/flecha_volver.asp" --></h4>      
	</hgroup>
    
    <mark>    
    <div class="tabla" align="center">
      <!-- Comentario 
        <p>
        <div class="titulo">COLUMNA #1</div>
        <div class="titulo">COLUMNA #2</div>
        <div class="titulo">COLUMNA #3</div>
        </p>
      -->

	<br>
	
	<%
	' Mostramos los datos del registro
	Response.Write RS.Fields(2) 'Contenido
	Response.Write "<BR>" & vbCrLf
	%>
	
	   <div id="mapa">
  <p>
  <div class="columna">
      <p><img src="/imagenes/<%Response.Write RS.Fields(5)%>" alt="<%Response.Write RS.Fields(5)%>" width="200" border="1" onmouseover="this.style.border='1px solid #128888';" onmouseout="this.style.border='1px solid #425364';" ></p>
    </div>
    <div class="columna">
      <p><img src="/imagenes/<%Response.Write RS.Fields(6)%>" alt="<%Response.Write RS.Fields(6)%>" width="200" border="1" onmouseover="this.style.border='1px solid #128888';" onmouseout="this.style.border='1px solid #425364';" ></p>
    </div>
    </p>
    
    <p>
    <div class="columna"></div>
    <div class="columna"></div>
    </p>

    </div>	
	
<%
	'listo...
End If

' Cerramos y limpiamos...
RS.Close
Set RS = Nothing
Conexion.Close
Set Conexion = Nothing

' Ahora mostramos los enlaces a las otras páginas con el resto de los registros...
If pagina_actual > 1 Then
	Response.Write "&nbsp;" & vbCrLf
End If

' mostramos la paginacion por numeros de página
For I = 1 To cant_paginas
	If I = pagina_actual Then
	Response.Write (I) 
	Else
	%> 
    <a href="/nota.asp?pagina=<%= I %>"><%= I %></a>
	  <%
	End If
Next 'I

'Fin...
%>
<br>
  

        
    </div>
    </mark>

  </article>
</section> 

  <aside id="derecha">
    <h4>Concreto</h4>
    <p>Trabajos industriales y particulares</p>
     <figure> &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
         <img src="/imagenes/hormigon_pulido_armado_llaneado.jpg" alt="Micropiso" width="100" border="1" align="center"> 
         <figcaption>Hormigon</figcaption>
     </figure> 
     <br>
     <p>Pavimentos en concreto asf&aacute;ltico - Pavimentos de hormig&oacute;n - Pisos de Hormig&oacute;n llaneado ferrocementado</p>     
  </aside>
  
  <footer><!--  Pie o Institucional -->
	<!-- #include virtual="/general/pie_institucional.asp" -->
  </footer>

  <!-- end .container --></div>
</body>
</html>

