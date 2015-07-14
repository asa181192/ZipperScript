	'Compresion de archivos mediante la aplicacion 7zip 
	'@Autor Alfredo Santiago Alvarado 

	'02/07/2015
	'****Actualizaciones****
	'05/07/2015- Agregado borrar archivos despues de zipear el archivo 
	'06/07/2015- Agregado zipeo archivo por archivo para evitar problemas al borrar archivos .
	
	
	centinelaPath = "C:\apps\PXDEL" 'Path del centinela verifica si los archivos fueron zipeados correctamente. 
    strPath = "\\DEV-UAT1\TransnetMQDB\basura" 'Path con los archivos origen a zipear 
	trashPath = "\\DEV-UAT1\TransnetMQDB\basura" 'Path con los archivos destino que fueron zipeados
	StrApp=  "C:\apps\PXDEL\7za.exe" 'Path de la aplicacion
    Set objshell = createobject("wscript.shell")
	Set objfso = createobject("scripting.filesystemobject")
	Set Objfolder=objfso.GetFolder(strPath)
	
	
	BolError = false 'indica cuando se genere un problema al momento de zipear archivos 
	
	IF objfso.FolderExists(centinelaPath&"\centinela\") THEN 'Verifica si existe la carpeta centinela , en caso de que no crea la carpeta centinela . 
	'0Set CarpetaCentinela = objfso.createfolder(strPath&"\centinela")'Se crea un archivo centinela que indica que todos los archivos fueron zipeados
		IF objfso.FileExists(centinelaPath&"\centinela\uat1.ctl") THEN
			 objfso.deleteFile centinelaPath&"\centinela\uat1.ctl"
		END IF 
	ELSE
		Set CarpetaCentinela = objfso.createfolder(centinelaPath&"\centinela")'Se crea un archivo centinela que indica que todos los archivos fueron zipeados
	END IF 
	
	
	For Each objfile In Objfolder.Files 'Itera sobre todos los archivos de la carpeta 

	if InStr(objfile.name,".txt")>0 THEN 'Se escpecifica la extension del archivo  
		RSpaces=Replace(objfile.name, " ", "")
		nombre = Split(RSpaces,".")
		filepathsource = strPath&"\"& objfile.name
		filepathdestiny = trashPath& "\" & nombre(0)
		comando = StrApp&" a -t7z "& filepathdestiny &" "& Chr(34)& filepathsource &Chr(34) &" -mx9"
		'Objshell.Exec(StrApp&" a -t7z "&nombre(0)&" "& Chr(34)& objfile.name &Chr(34)&" -mx9")
		'MsgBox filepath
			On Error Resume Next
			objshell.Run comando,0,True 'Corre el comando para llamar al programa y  realizar el zipeo de archivos 
			If Err.Number <> 0 THEN
				BolError = true 
				Else
				objfso.deletefile objfile 'Borra el archivo que fue zipeado 
				
			End IF
	END IF 
	Next 
	
	
	IF BolError=false THEN ' Crea el archivo centinela dentro de la carpeta centinela . 
	Set archivotexto = objfso.createtextfile(centinelaPath&"\centinela"&"\uat1.ctl",true)
	archivotexto.writeline "Archivos Zipeados correctamente"
	archivotexto.close	
	END IF 
	
	
	