'Programa: Worklists_Concatena.vbs
'Objetivo: Transforma os arquivos gerados pela leitura dos WORKLISTS de usuários em um só arquivo
'Data: 04/10/200148


dim x,arrStr,WorkList, strText, oFile
dim fs,objReadFile, ObjInputFile, ObjOutputFile
Const ForReading = 1

set fs=CreateObject("Scripting.FileSystemObject")
set objReadFile = fs.OpenTextFile("C:\Users\GLB00156\Documents\SAP\SAP GUI\INPUT_WORKLISTS.txt")


Do while NOT objReadFile.AtEndOfStream
	arrStr         = Split(objReadFile.ReadLine,";")
	WorkList       = arrStr(0)
  oFile = "C:\Users\GLB00156\Documents\SAP\SAP GUI\" & WorkList

'//OPEN FILE and READ
Set objFileToRead = CreateObject("Scripting.FileSystemObject").OpenTextFile(oFile,1)
strFileText = strFileText  & vbCrLf & objFileToRead.ReadAll()
objFileToRead.Close


' ///PASTE 
Set objFSO = CreateObject("Scripting.FileSystemObject") 
Set objFileToWrite = objFSO.OpenTextFile("C:\Users\GLB00156\Documents\SAP\SAP GUI\ALL_WORKLIST.txt", 2)
objFileToWrite.Write strFileText
objFileToWrite.Close

Loop

objReadFile.Close
set objReadFile = Nothing
set js = Nothing
set arrStr = nothing

Msgbox "Pronto. Acabei !"