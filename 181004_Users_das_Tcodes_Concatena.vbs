'Programa: Users_das_Tcodes_Concatena.vbs
'Objetivo: Transforma os arquivos gerados pela leitura dos usu�rios de cada transa��o em um s� arquivo
'Data: 04/10/200148


dim x,arrStr,TCode, strText, oFile
dim fs,objReadFile, ObjInputFile, ObjOutputFile
Const ForReading = 1

set fs=CreateObject("Scripting.FileSystemObject")
set objReadFile = fs.OpenTextFile("C:\Users\GLB00156\Documents\SAP\SAP GUI\181205_COLETA USERS POR TCODE_JA USADA\0_Transacoes_SAP.csv")


Do while NOT objReadFile.AtEndOfStream
	arrStr         = Split(objReadFile.ReadLine,";")
	TCode	       	 = arrStr(0)
  oFile = "C:\Users\GLB00156\Documents\SAP\SAP GUI\181205_COLETA USERS POR TCODE_JA USADA\" & TCode

'//OPEN FILE and READ
Set objFileToRead = CreateObject("Scripting.FileSystemObject").OpenTextFile(oFile,1)
strFileText = strFileText  & vbCrLf & objFileToRead.ReadAll()
objFileToRead.Close


' ///PASTE 
Set objFSO = CreateObject("Scripting.FileSystemObject") 
Set objFileToWrite = objFSO.OpenTextFile("C:\Users\GLB00156\Documents\SAP\SAP GUI\181205_COLETA USERS POR TCODE_JA USADA\All_Tcodes_New.txt", 2)
objFileToWrite.Write strFileText
objFileToWrite.Close

Loop

objReadFile.Close
set objReadFile = Nothing
set js = Nothing
set arrStr = nothing

Msgbox "Pronto. Acabei !"