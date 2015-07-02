	'Compresion de archivos mediante la aplicacion 7zip 
	'@Autor Alfredo Santiago Alvarado 

	'02/07/2015
    strPath = "\\ASANTIAGO-PC1\CursosRedesSeguridadProgramacion\TEST" 'Path con los archivos a zipear 
	StrApp=  "\\ASANTIAGO-PC1\CursosRedesSeguridadProgramacion\APPS\7za.exe" 'Path de la aplicacion
    Set objshell = createobject("wscript.shell")
	Set objfso = createobject("scripting.filesystemobject")
	Set Objfolder=objfso.GetFolder(strPath)
	
	For Each objfile In Objfolder.Files
	if InStr(objfile.name,".txt")>0 THEN
	RSpaces=Replace(objfile.name, " ", "")
	nombre = Split(RSpaces,".")
	Objshell.Exec(StrApp&" a -t7z "&nombre(0)&" "& Chr(34)& objfile.name &Chr(34))'Se debe indicar el path de lrecurso con la aplicacion

	
	END IF 
	Next 

 

	
