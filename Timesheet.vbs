'******************************************************************************
'AUTO-TIMESHEET BASEADO NOS EVENTOS DE WINLOGON E LOGOFF DO WINDOWS
'PREPARA UMA BASE DE DADOS QUE � UTILIZADA PELO EXCEL DE APONTAMENTOS
'AUTOMATICAMENTE LAN�ANDO OS APONTAMENTOS BASEADOS NESTES EVENTOS


'******************************************************************************
'Vari�veis
'objExcel / objworksheet = C:\VBScript\Timesheet.xlsx
'intLinUltimo     = �ltima linha lan�ada na coluna B
'dtUltimo         = �ltima data lan�ada na coluna A
'dtUltimox        = Convers�o de dtUltimo por CDate (para eventos 6009)
'dtUltimoy	  = Convers�o de dt�ltimo por CDate (para eventos 6006)
'objSWbemServices = Vari�vel de acesso ao WMI
'colTimeZone      = Cole��o da query de Time Zone
'objTimeZone      = Objeto Time Zone
'strBias          = String de apoio para forma��o da data ISO
'dtmCurrentDate   = DtUltimox
'dtmDay           = Dia da data (testa se � menor que 10 e coloca um "0"
'dtmMonth         = M�s da data (testa se � menor que 10 e coloca um "0"
'dtmTargetDate    = String da data ISO final 
'objWMIService    = Vari�vel de acesso ao WMI
'colEvents        = Cole��o de Eventos de login e logoff
'n                = Contador de apoio � matriz de eventos
'intOcorr	  = N�mero de ocorr�ncias de eventos
'intItem          = Item de datahora
'arrEvent()       = Matriz de eventos de datahora




'******************************************************************************
'Abre o excel dos apontamentos

Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = False
Set objWorksheet = objExcel.Workbooks.Open("C:\VBScript\Timesheet.xlsx")



'******************************************************************************
'Qual o �ltimo apontamento e sua data ?
'
set intLinUltimo = GetObject(,"Excel.application")
set dtUltimo = GetObject (,"Excel.application")
intLinhas = 1
	Do until len(intLinUltimo.cells(intLinhas, 2).value) = 0
     	intLinhas = intLinhas + 1
	Loop
	DtUltimo = DtUltimo.cells (intLinhas,1).value
	DtUltimox = CDate (DtUltimo)
	DtUltimoy = DateAdd ("d", 1, DtUltimox)
'	Wscript.Echo "A �ltima linha de lan�amento � a " & intLinhas
'	Wscript.Echo "A pr�xima data de lan�amento � " & dtUltimox


'******************************************************************************
'Essa rotina cria uma ISO date correta, a partir de DtUltimox Ex: 17/05/2012 = 20120517000000.000000 -000 

strComputer = "."
Set objSWbemServices = GetObject("winmgmts:" _
 & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colTimeZone = objSWbemServices.ExecQuery _
 ("SELECT * FROM Win32_TimeZone")
For Each objTimeZone in colTimeZone
 strBias = objTimeZone.Bias
Next

dtmCurrentDatex = DtUltimox
dtmTargetDatex = Year(dtmCurrentDatex)

dtmMonthx = Month(dtmCurrentDatex)
If Len(dtmMonthx) = 1 Then
 dtmMonthx = "0" & dtmMonthx
End If

dtmTargetDatex = dtmTargetDatex & dtmMonthx

dtmDayx = Day(dtmCurrentDatex)
If Len(dtmDayx) = 1 Then
 dtmDayx = "0" & dtmDayx
End If

dtmCurrentDatey = DtUltimoy
dtmTargetDatey = Year(dtmCurrentDatey)

dtmMonthy = Month(dtmCurrentDatey)
If Len(dtmMonthy) = 1 Then
 dtmMonthy = "0" & dtmMonthy
End If

dtmTargetDatey = dtmTargetDatey & dtmMonthy

dtmDayy = Day(dtmCurrentDatey)
If Len(dtmDayy) = 1 Then
 dtmDayy = "0" & dtmDayy
End If


dtmTargetDatex = dtmTargetDatex & dtmDayx & "000000.000000"
dtmTargetDatex = dtmTargetDatex & Cstr(strBias)

dtmTargetDatey = dtmTargetDatey & dtmDayy & "000000.000000"
dtmTargetDatey = dtmTargetDatey & Cstr(strBias)


'*********************************************************************************************
'Essa rotina busca os eventos em DtUltimox (ou dtmTargetDate) e seleciona o winlogon mais cedo
'

Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate,(Security)}!\\" _
    & strComputer & "\root\cimv2")

Set colEvents = objWMIService.ExecQuery _
    ("Select * from Win32_NTLogEvent Where Logfile = 'System' and EventCode = '6009' and TimeWritten >= '" & dtmTargetDatex & "'")

Dim n
Dim intOcorr
Dim arrEvent ()
intOcorr = colEvents.count
ReDim arrEvent (intOcorr)
For each objEvent in colEvents
intItem = objEvent.Timewritten
arrEvent(n) = WMIDateStringToDate (intItem)
n=n+1
Next
LoginMaisCedo(arrEvent)
n=0
Set intOcorr=nothing
Erase arrEvent



'*********************************************************************************************
'Essa rotina busca os eventos em DtUltimox (ou dtmTargetDate) e seleciona o winloff mais tarde
'

Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate,(Security)}!\\" _
    & strComputer & "\root\cimv2")

Set colEvents = objWMIService.ExecQuery _
    ("Select * from Win32_NTLogEvent Where Logfile = 'System' and EventCode = '6006' and TimeWritten <= '" & dtmTargetDatey & "'")

intOcorr = colEvents.count
ReDim arrEvent (intOcorr)
For each objEvent in colEvents
intItem = objEvent.Timewritten
arrEvent(n) = WMIDateStringToDate (intItem)
n=n+1
Next
LogoffMaisTarde(arrEvent)
n=0
Set intOcorr=nothing
Erase arrEvent

objExcel.Workbooks(1).Save 'Save the workbook, not excel file
objExcel.Workbooks(1).Close 'Close the workbook then finally quit excel
objExcel.Quit 

Wscript.Echo "Alexandre, timesheet para "& DtUltimox & " atualizado."



'*********************************************************************************************
'Fun��o converte data de formato 20120517000000.000000 -000 para formato dd/mm/aaaa hh:mm:ss
'
Function WMIDateStringToDate(intItem)
    	
	WMIDateStringToDate = CDate(Mid(intItem, 7, 2) & "/" & _
        Mid(intItem, 5, 2) & "/" & Left(intItem, 4) _
            & " " & Mid (intItem, 9, 2) & ":" & _
                Mid(intItem, 11, 2) & ":" & Mid(intItem, _
                    13, 2))
	WMIDateStringToDate = DateAdd ("h", -2, WMIDateStringToDate) 'Acerta o hor�rio de Brasilia (-3) ou ver�o (-2)
	
End Function




'*********************************************************************************************
'Fun��o seleciona o evento Logon mais cedo da data a ser lan�ada 
'

Function LoginMaisCedo(AnArray())
Dim MinItem
Dim Item
MinItem=AnArray(0)
   For each Item in AnArray
		intCheckDT = DateDiff("yyyy",MinItem,Item)*12441600 _
		 	   + DateDiff("m",MinItem,Item)*1036800 _
			   + DateDiff("d",MinItem,Item)*86400 _
			   + DateDiff("h",MinItem,Item)*3600 _
			   + DateDiff("n",MinItem,Item)*60 _
			   + DateDiff("s",MinItem,Item)
	If Item <> 0 AND intCheckDT < 0 Then 
	MinItem = Item
	End If
   Next
objExcel.Cells(intLinhas, 2) = MinItem
End Function



'*********************************************************************************************
'Fun��o seleciona o evento Logoff mais tardio da data a ser lan�ada
'

Function LogoffMaisTarde(AnArray())
Dim MaxItem
Dim Item
MaxItem=AnArray(0)
   For e = 0 to intOcorr
	Item=AnArray(e)
		intCheckDT = DateDiff("yyyy",MaxItem,Item)*12441600 _
		   	   + DateDiff("m",MaxItem,Item)*1036800 _
			   + DateDiff("d",MaxItem,Item)*86400 _
			   + DateDiff("h",MaxItem,Item)*3600 _
			   + DateDiff("n",MaxItem,Item)*60 _
			   + DateDiff("s",MaxItem,Item)
	If Item <> 0 AND intCheckDT > 0 Then 
	MaxItem = Item
	End If
   Next
objExcel.Cells(intLinhas, 5) = MaxItem
End Function

