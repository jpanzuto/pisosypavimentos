<%
' OJO tiene info si viene de un click unicamente
url_atra = request.servervariables("HTTP_REFERER")

tmp = InStr(1, Request.ServerVariables("HTTP_REFERER"), "pisosypavimentos.com.ar")
If tmp > 0 Then url_atras = request.servervariables("HTTP_REFERER") else url_atras = "/servicios.asp" End If
%>

<a href="<% =url_atras %>"><img src="/img-general/flecha-volver.png" alt="volver" onmouseover="this.src='/img-general/flecha-volver-seleccion.png'" onmouseout="this.src='/img-general/flecha-volver.png'"></a> 