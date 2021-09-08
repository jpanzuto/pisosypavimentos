<script language="VBScript" runat=server>



Const adOpenForwardOnly = 0
Const adLockReadOnly = 1
Const adCmdText = &H0001
Const adUseClient = 3
Const adLockOptimistic = 3

'DATOS SERVIDOR
Const servidor = "servidor"               'SERVIDOR DE BASE DE DATOS
Const baseDatos = "basedatos"			  'NOMBRE DE LA BASE DE DATOS
	
Dim oConn_

FUNCTION HTMLSafe (ByVal strData)
	strData = Replace(strData, "&", "&amp;")
    strData = Replace(strData, "'", "&#39;")
    strData = Replace(strData, "<", "&lt;")
    strData = Replace(strData, ">", "&gt;")
	strData = Replace(strData, chr(13), "<br>")
    HTMLSafe = strData
END FUNCTION

FUNCTION SQLSafe(ByVal str)
    SQLSafe = Replace(str, "'", "''")
END FUNCTION

FUNCTION Quote(ByVal str)
    Quote = "'" & str & "'"
END FUNCTION

'Se conecta con el servidor de base de datos
FUNCTION Establecer_Conexion (ByVal usr, ByVal pass)
	Set oConn_ = CreateObject("ADODB.Connection")
	'Utilizamos un servidor de bases de datos
	'oConn_.Open "Provider=SQLOLEDB; Data Source="&servidor&"; Initial Catalog="&baseDatos&"; " & "User Id="&usr&"; Password="&pass&";"
	
	'Utilizamos MS Access
	oConn_.Open("DRIVER={Microsoft Access Driver (*.mdb)}; DBQ=" & Server.MapPath("/general/base.mdb"))
END FUNCTION


'Ejecuta el procedimiento almacenado indicado.
FUNCTION Ejecutar_Procedimiento (ByVal procedimiento)
	Dim cmd_
	Set cmd_ = CreateObject("ADODB.Command")
	
	cmd_.ActiveConnection = oConn_
	cmd_.CommandText = procedimiento
	
	Set Ejecutar_Procedimiento = cmd_.Execute
END FUNCTION


'Devuelve un recorset con los resultados de la consulta
FUNCTION Consultar_Registros (ByVal tabla, ByVal Consulta)
	Dim oRS_
	Dim SQL_

	Set oRS_ = CreateObject("ADODB.RecordSet")	
	
	IF IsEmpty(Consulta) then
		SQL_ = "select * from " & tabla
	ELSE
		SQL_ = "select * from " & tabla & " where " & Consulta
	END IF
	
	oRS_.Open SQL_, oConn_, adOpenForwardOnly, adLockOptimistic, adCmdText
	
	Set Consultar_Registros = oRS_
	set oRS_ = Nothing
END FUNCTION

'Devuelve un recorset con los resultados de la consulta libre
FUNCTION Consulta (ByVal strConsulta)
	Dim oRS_
	Set oRS_ = CreateObject("ADODB.RecordSet")	
	oRS_.Open strConsulta, oConn_, adOpenForwardOnly, adLockOptimistic, adCmdText
	
	Set Consulta = oRS_
	set oRS_ = Nothing
END FUNCTION

'Devuelve un recorset con los resultados de la consulta
FUNCTION Consultar (ByVal tabla, ByVal Registros, ByVal Consulta)
	Dim oRS_
	Dim SQL_

	Set oRS_ = CreateObject("ADODB.RecordSet")	
	
	IF IsEmpty(Consulta) then
		SQL_ = "select " & Registros & " from " & tabla
	ELSE
		SQL_ = "select " & Registros & " from " & tabla & " where " & Consulta
	END IF
	
	oRS_.Open SQL_, oConn_, adOpenForwardOnly, adLockOptimistic, adCmdText
	
	Set Consultar = oRS_
END FUNCTION

'Devuelve un recorset con los resultados de la consulta y lo deja ABIERTO para poder modificar cosas
'ATENCIÓN: se debe cerrar el recorset desde donde se llama
FUNCTION ConsultarYModificar (ByVal tabla, ByVal Registros, ByVal Consulta)
	Dim oRS_
	Dim SQL_

	Set oRS_ = CreateObject("ADODB.RecordSet")	
	
	IF IsEmpty(Consulta) then
		SQL_ = "select " & Registros & " from " & tabla
	ELSE
		SQL_ = "select " & Registros & " from " & tabla & " where " & Consulta
	END IF
	
	oRS_.Open SQL_, oConn_, adOpenForwardOnly, adLockOptimistic, adCmdText
	
	Set ConsultarYModificar = oRS_
END FUNCTION

'Cierra la conexión con la base de datos
FUNCTION Cerrar_Conexion
	oConn_.Close
	Set oConn_ = Nothing
END FUNCTION

'Inserta un registro en la base de datos
FUNCTION Insertar_Registro (ByVal tabla, ByVal arrDatos)
	Dim oRS_
	Set oRS_ = CreateObject("ADODB.RecordSet")
	i = 0
	lon = UBound(arrDatos)
	str1 = ""
	str2 = ""
	coma = ","
	for each str in arrDatos
		if (i Mod 2) = 0 Then
			IF (lon - 2) = i THEN
				coma = ""
			END IF
			str1 = str1 & str & coma & " "
		Else
			IF (lon - 1) = i THEN
				coma = ""
			END IF
			str2 = str2 & str & coma & " "
		End If
		i = i + 1
	Next
	SQL = "insert into " & tabla & " (" & str1 & ") values (" & str2 & ")"
	oRS_.Open SQL, oConn_, adOpenForwardOnly, adLockOptimistic, adCmdText
	set oRS_ = Nothing
END FUNCTION

'Modificar registros en la base de datos
FUNCTION Modificar_Registro (ByVal tabla, ByVal arrDatos, ByVal consulta)
	Dim oRS_
	Set oRS_ = CreateObject("ADODB.RecordSet")
	i = 0
	lon = UBound(arrDatos)
	igual = " = "
	str1 = ""
	coma = ","
	for each str in arrDatos
		if (i Mod 2) = 0 Then
			IF (lon - 2) = i THEN
				'igual = ""
			END IF
			str1 = str1 & str
		Else
			IF (lon - 1) = i THEN
				coma = ""
				'igual = ""
			END IF
			str1 = str1 & igual & str & coma & " "
		End If
		i = i + 1
	Next
	SQL = "update " & tabla & " set " & str1 & " where (" & consulta & ")"

	oRS_.Open SQL, oConn_, adOpenForwardOnly, adLockOptimistic, adCmdText
	set oRS_ = Nothing
END FUNCTION

'Borra los registros de la tabla que cumplan la condición de la consulta
FUNCTION Borrar_Registros (ByVal tabla, ByVal Consulta)
	Dim oRS_
	Dim SqlAux
	Dim SQL
	
	Set oRS_ = CreateObject("ADODB.RecordSet")
	
	IF IsEmpty(Consulta) then
		SqlAux = ""
	ELSE
		SqlAux = " where " & SQLSafe(Consulta)
	END IF
	
	SQL = "select * from " & tabla & SqlAux
	oRS_.Open SQL, oConn_, adOpenForwardOnly, adLockOptimistic, adCmdText
	
	If (not oRS_.EOF) Then 'Si hay algo que borrar
		oRS_.first
		while (not oRS_.EOF)
			oRS_.Delete
			oRS_.Update
		wend
	End If
	
	set oRS_ = Nothing
END FUNCTION
</script>
