<%
'creamos un array para contener los nombres de las imagenes
'tan grande (al menos) como el numero de imagenes que contenga el directorio
dim x(200)

'Carpeta que contiene nuestras imagenes

Const mypath="\imagenes"

'creacion del objeto FSO
Set filesystem = CreateObject("Scripting.FileSystemObject")
Set folder = filesystem.GetFolder(server.mappath(mypath))
Set filecollection = folder.Files

'carga del array 
idx=0

For Each file in filecollection
idx=idx+1
x(idx)=file.name
Next

'Elegimos una imagen al azar
randomize timer
whichNo=int(rnd()*idx)+1

'Destruimos ls objetos
set filesystem=nothing
set folder=nothing
set filecollection=nothing

'Mostramos la imagen seleccionada
response.write "<img src=" & mypath & "/"
response.write x(whichNO)& " alt=" & x(whichNo) & " width=100 border=1>"
%>