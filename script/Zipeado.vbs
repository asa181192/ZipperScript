	'Compresion de archivos mediante la aplicacion 7zip 
	'@Autor Alfredo Santiago Alvarado 

	'02/07/2015
	'****Actualizaciones****
	'05/07/2015- Agregado borrar archivos despues de zipear el archivo 
	'06/07/2015- Agregado zipeo archivo por archivo para evitar problemas al borrar archivos .
	
	
    strPath = "\\ASANTIAGO-PC1\CursosRedesSeguridadProgramacion\TEST" 'Path con los archivos a zipear 
	StrApp=  "\\ASANTIAGO-PC1\CursosRedesSeguridadProgramacion\APPS\7za.exe" 'Path de la aplicacion
    Set objshell = createobject("wscript.shell")
	Set objfso = createobject("scripting.filesystemobject")
	Set Objfolder=objfso.GetFolder(strPath)
	
	BolError = false 
	
	For Each objfile In Objfolder.Files

	if InStr(objfile.name,".txt")>0 THEN
		RSpaces=Replace(objfile.name, " ", "")
		nombre = Split(RSpaces,".")
		comando = StrApp&" a -t7z "&nombre(0)&" "& Chr(34)& objfile.name &Chr(34)&" -mx9"
		'Objshell.Exec(StrApp&" a -t7z "&nombre(0)&" "& Chr(34)& objfile.name &Chr(34)&" -mx9")
			On Error Resume Next
			objshell.Run comando,0,True
			If Err.Number <> 0 THEN
				BolError = true 
				Else
				objfso.deletefile objfile
				
			End IF
	END IF 
	Next 
	
	IF objfso.FolderExists(strPath&"\centinela\") THEN
	'0Set CarpetaCentinela = objfso.createfolder(strPath&"\centinela")'Se crea un archivo centinela que indica que todos los archivos fueron zipeados
	
	Else
		Set CarpetaCentinela = objfso.createfolder(strPath&"\centinela")'Se crea un archivo centinela que indica que todos los archivos fueron zipeados
	END IF 
	
	IF BolError=false THEN
	Set archivotexto = objfso.createtextfile(strPath&"\centinela"&"\centinela.txt",true)
	archivotexto.writeline "Archivos Zipeados correctamente"
	archivotexto.close	
	END IF 
	
	
	