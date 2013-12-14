* \dev\ilib_my.prg
* 17/06/13 Se creo ete archivo y trajeron 4 rutinas Wsh_* de ilib
* 19/04/09 Alex.  Se añadieron 4 rutinas que llaman a WindowsScriptingHost
FUNCTION Wsh_MapNetworkDrive
LPARAMETERS tcNombreLocal,tcNombreRemoto,tlActualizarPerfil,tcUser,tcPassword
LOCAL loNet AS WScript.Network,llStatus
loNet = CreateObject("WScript.Network")
TRY
	loNet.MapNetworkDrive(tcNombreLocal,tcNombreremoto)
	llStatus = .T.
CATCH
ENDTRY
RETURN llStatus

FUNCTION Wsh_RemoveNetworkDrive
LPARAMETERS tcNombreLocal
LOCAL loNet AS WScript.Network,llStatus
loNet = CreateObject("WScript.Network")
TRY
	loNet.RemoveNetworkDrive(tcNombreLocal)
	llStatus = .T.
CATCH
ENDTRY
RETURN llStatus

* Retorna el objeto devuelto por EnumNetworkDrives
* Si recibe un arreglo como parametro lo llena con los valores regresados por EnumNetworkDrives
FUNCTION Wsh_EnumNetworkDrives
LPARAMETERS taDrives
EXTERNAL ARRAY taDrives
*!*	Set WshNetwork = WScript.CreateObject("WScript.Network")
*!*	Set oDrives = WshNetwork.EnumNetworkDrives
*!*	Set oPrinters = WshNetwork.EnumPrinterConnections
*!*	WScript.Echo "Network drive mappings:"
*!*	For i = 0 to oDrives.Count - 1 Step 2
*!*		WScript.Echo "Drive " & oDrives.Item(i) & " = " & oDrives.Item(i+1)
*!*	Next
*!*	WScript.Echo 
*!*	WScript.Echo "Network printer mappings:"
*!*	For i = 0 to oPrinters.Count - 1 Step 2
*!*		WScript.Echo "Port " & oPrinters.Item(i) & " = " & oPrinters.Item(i+1)
*!*	Next
LOCAL loNet AS WScript.Network,loDrives,i
loNet = CreateObject("WScript.Network")
loDrives = loNet.EnumNetworkDrives()
IF PARAMETERS() > 0
	DIMENSION taDrives[loDrives.Count / 2,2]
	FOR i = 1 TO loDrives.Count / 2
		taDrives[i,1] = loDrives.Item[i * 2 - 2 ]
		taDrives[i,2] = loDrives.Item[i * 2 - 1]
	ENDFOR
ENDIF
RETURN loDrives

FUNCTION Wsh_Run
LPARAMETERS tcCommand,tcWindowStyle,tlWaitOnReturn
*!*	Valores de tcWindowStyle:
*!*	0 Hides the window and activates another window.
*!*	1 Activates and displays a window. If the window is minimized or maximized, the system restores it to its original size
*!*	  and position. An application should specify this flag when displaying the window for the first time.
*!*	2 Activates the window and displays it as a minimized window. 
*!*	3 Activates the window and displays it as a maximized window. 
*!*	4 Displays a window in its most recent size and position. The active window remains active.
*!*	5 Activates the window and displays it in its current size and position.
*!*	6 Minimizes the specified window and activates the next top-level window in the Z order.
*!*	7 Displays the window as a minimized window. The active window remains active.
*!*	8 Displays the window in its current state. The active window remains active.
*!*	9 Activates and displays the window. If the window is minimized or maximized, the system restores it to its original size
*!*	  and position. An application should specify this flag when restoring a minimized window.
*!*	10 Sets the show-state based on the state of the program that started the application.
*Dim oShell
*Set oShell = WScript.CreateObject ("WScript.Shell")
*loShell.run "cmd.exe /K CD C:\ & Dir"
*Set oShell = Nothing
LOCAL loShell AS WScript.Shell,lnRetVal
loShell = CREATEOBJECT("WScript.Shell")
lnRetVal = loShell.Run(tcCommand,tcWindowStyle,tlWaitOnReturn)
DECLARE Sleep IN kernel32 INTEGER dwMilliseconds
Sleep(50)
RETURN lnRetVal


* 17/06/13 Se crearon rutinas para cambiar formato de Drive de \\Server\Share a X: y viceversa
* Recibe Path (\\server\share\Directory\File o X:\Directory\File) y devuelve drive (\\Server\Share o x:)
FUNCTION Path2Drive
LPARAMETERS tcPath
	LOCAL lcDrive
	tcPath = ALLTRIM(tcPath)
	DO CASE
		CASE LEFT(tcPath,2) = '\\'
			lcDrive = LEFT(tcPath+'\',AT('\',tcPath+'\',4)-1)
		CASE SUBSTR(tcPath,2,1) = ':'
			lcDrive = LEFT(tcPath,2)
		OTHERWISE
			lcDrive = ''
	ENDCASE
RETURN lcDrive

* Recibe RemoteDrive (\\Server\Share) y devuelve LocalDrive (X:)
FUNCTION RemoteDrive2LocalDrive
LPARAMETERS tcDrive
	tcDrive = ALLTRIM(UPPER(tcDrive))
	LOCAL laDrives[1,2],i,lcLocalDrive
	Wsh_EnumNetworkDrives(@laDrives)
	lcLocalDrive = ''
	FOR i = 1 TO ALEN(laDrives,1)
		IF UPPER(laDrives[i,2]) = tcDrive
			lcLocalDrive = laDrives[i,1]
		ENDIF
	ENDFOR
RETURN lcLocalDrive

* Recibe LocalDrive (X:) y devuelve RemoteDrive (\\Server\Share)
FUNCTION LocalDrive2RemoteDrive
LPARAMETERS tcDrive
	tcDrive = ALLTRIM(UPPER(tcDrive))
	LOCAL laDrives[1,2],i,lcRemoteDrive
	Wsh_EnumNetworkDrives(@laDrives)
	lcRemoteDrive = ''
	FOR i = 1 TO ALEN(laDrives,1)
		IF UPPER(laDrives[i,1]) = tcDrive
			lcRemoteDrive = laDrives[i,2]
		ENDIF
	ENDFOR
RETURN lcRemoteDrive

* Recibe Path (\\server\share\Directory\File o X:\Directory\File) y devuelve LocalDrive (x:)
FUNCTION Path2LocalDrive
LPARAMETERS tcPath
	LOCAL lcRetVal
	tcPath = ALLTRIM(tcPath)
	lcDrive = Path2Drive(tcPath)
	DO CASE
		CASE LEFT(lcDrive,2) = '\\'
			lcRetVal = RemoteDrive2LocalDrive(lcDrive)
		CASE SUBSTR(lcDrive,2,1) = ':'
			lcRetVal = lcDrive
		OTHERWISE
			lcRetVal = ''
	ENDCASE
RETURN lcRetVal

* Recibe Path (\\server\share\Directory\File o X:\Directory\File) y devuelve RemoteDrive (\\server\share)
FUNCTION Path2RemoteDrive
LPARAMETERS tcPath
	LOCAL lcDrive,lcRetVal
	tcPath = ALLTRIM(tcPath)
	lcDrive = Path2Drive(tcPath)
	DO CASE
		CASE LEFT(lcDrive,2) = '\\'
			lcRetVal = lcDrive
		CASE SUBSTR(lcDrive,2,1) = ':'
			lcRetVal = LocalDrive2RemoteDrive(lcDrive)
		OTHERWISE
			lcRetVal = ''
	ENDCASE
RETURN lcRetVal

FUNCTION ForceDrive
LPARAMETERS tcPath,tcDrive
	LOCAL lcRemoteDrive,lcLocalDrive,lcRetVal
	DO CASE
		CASE LEFT(tcPath,2) = '\\'
			lcRemoteDrive = Path2RemoteDrive(tcPath)
			lcLocalDrive = RemoteDrive2LocalDrive(lcRemoteDrive)
			lcRetVal = lcLocalDrive + SUBSTR(tcPath,LEN(lcRemoteDrive)+1)
		CASE SUBSTR(tcPath,2,1) = ':'
			lcLocalDrive = Path2LocalDrive(tcPath)
			lcRemoteDrive = LocalDrive2RemoteDrive(lcLocalDrive)
			lcRetVal = lcRemoteDrive + SUBSTR(tcPath,LEN(lcLocalDrive)+1)
		OTHERWISE 
	ENDCASE
RETURN lcRetVal

FUNCTION ForceLocalDrive
LPARAMETERS tcPath
	LOCAL lcRemoteDrive,lcLocalDrive,lcRetVal
	DO CASE
		CASE LEFT(tcPath,2) = '\\'
			lcRemoteDrive = Path2RemoteDrive(tcPath)
			lcLocalDrive = RemoteDrive2LocalDrive(lcRemoteDrive)
			lcRetVal = lcLocalDrive + SUBSTR(tcPath,LEN(lcRemoteDrive)+1)
		CASE SUBSTR(tcPath,2,1) = ':'
			lcRetVal = tcPath
		OTHERWISE 
	ENDCASE
RETURN lcRetVal

FUNCTION ForceRemoteDrive
LPARAMETERS tcPath
	LOCAL lcRemoteDrive,lcLocalDrive,lcRetVal
	DO CASE
		CASE LEFT(tcPath,2) = '\\'
			lcRetVal = tcPath
		CASE SUBSTR(tcPath,2,1) = ':'
			lcLocalDrive = Path2LocalDrive(tcPath)
			lcRemoteDrive = LocalDrive2RemoteDrive(lcLocalDrive)
			lcRetVal = lcRemoteDrive + SUBSTR(tcPath,LEN(lcLocalDrive)+1)
		OTHERWISE 
	ENDCASE
RETURN lcRetVal


FUNCTION FSO_GetAbsolutePathName(tcPath)
LOCAL loFSO AS Scripting.FileSystemObject
loFSO = CreateObject("Scripting.FileSystemObject")
RETURN loFSO.GetAbsolutePathName(tcPath)


PROCEDURE MyDebugOut
LPARAMETERS tcMensaje,tuContador,tnDestino
* MyDebugOut - Graba log con mensaje indicado y tiempo transcurrido en cuatro contadores de tiempo
* tcMensaje  = Texto del mensaje
* tuContador = Numero del contador de tiempo a mostrar
*			   Si se suma 100 al contador éste se resetea antes de mostrar el tiempo transcurrido
*			   Se puede mostrar mas de un contador enviandolos entre comillas y separados por comas. Por ejemplo: "2,3"
* tnDestino  = 1 o .F.	El mensaje se graba y se muestra en DebugOut
*			 = 2		El mensaje solo se muestra en ventana DebugOut
*			 = 3		El mensaje solo se graba
*			   Si se suma 100 al destino se imprime '----------------------' antes de emitir el mensaje
*			   Si se suma 1000 al destino se borra el archivo DevLog antes de emitir el mensaje

* Setup
	LOCAL lnSeconds,lcDevLog,lcMensaje,i,lnContador
	LOCAL laMencion[1]	&& Arreglo de contadores mencionados en tuContador, en el orden en que fueron mencionados
	LOCAL laResetear[4]	&& Arreglo que indica si usuario solicito resetear contador i

* Parametros
	tuContador = IIF(EMPTY(tuContador),1,tuContador)
	tnDestino  = IIF(EMPTY(tnDestino),1,tnDestino)

* Creacion de arreglo de contadores
	lnSeconds = SECONDS()
	IF TYPE('_SCREEN.GrabarLog_Seconds[4]') = 'U'
		_SCREEN.AddProperty('GrabarLog_Seconds[4]',lnSeconds)
	ENDIF
	
* Guardamos contadores que usuario menciono y reseteamos los que solicito resetear
	IF VARTYPE(tuContador) = 'N'
		IF tuContador < 100
			laMencion[1] = tuContador
		ELSE
			laMencion[1]		     = tuContador - 100
			laResetear[laMencion[1]] = .T.
			_SCREEN.GrabarLog_Seconds[laMencion[1]] = lnSeconds
		ENDIF
	ENDIF
	IF VARTYPE(tuContador) = 'C'
		FOR i = 1 TO ALINES(laMencion,tuContador,1,',')
			laMencion[i] = INT(VAL(laMencion[i]))
			IF laMencion[i] > 100
				laMencion[i]			 = laMencion[i] - 100
				laResetear[laMencion[i]] = .T.
				_SCREEN.GrabarLog_Seconds[laMencion[i]] = lnSeconds
			ENDIF
		ENDFOR
	ENDIF

* DevLog
	IF TYPE('_SCREEN.GrabarLog_DevLog') = 'U'
		_SCREEN.AddProperty('GrabarLog_DevLog','DEVLOG_'+GETENV("COMPUTERNAME")+'.log')
	ENDIF
	IF tnDestino > 1000
		tnDestino = tnDestino - 1000
		DEBUGOUT ' '
		STRTOFILE('',_SCREEN.GrabarLog_DevLog)
	ENDIF
	IF tnDestino > 100
		tnDestino = tnDestino - 100
		DEBUGOUT ' '
		STRTOFILE(REPLICATE('-',40) + CHR(13) + CHR(10),_Screen.GrabarLog_DevLog,.T.)
	ENDIF

* Mensaje
	DO CASE
		CASE LEN(tcMensaje) < 60
			tcMensaje = PADR(tcMensaje,60)
		CASE LEN(tcMensaje) < 80
			tcMensaje = PADR(tcMensaje,80)
		OTHERWISE 
			tcMensaje = PADR(tcMensaje,100)
	ENDCASE
	lcMensaje = tcMensaje + ' Segundos: ' + PADR(TRANSFORM(lnSeconds),10) + ' Transcurrido: '
	FOR lnContador = 1 TO 4
		lcMensaje = lcMensaje + IIF(ASCAN(laMencion,lnContador) > 0, ;
									PADR(TRANSFORM(lnSeconds - _SCREEN.GrabarLog_Seconds[lnContador],'9999.999'),20), ;
									REPLICATE(' ',20))
		* Reseteamos contadores mencionados en tuContador que no fueron reseteados antes de generar mensaje
		IF ASCAN(laMencion,lnContador) > 0 AND !laResetear[lnContador]
			_SCREEN.GrabarLog_Seconds[lnContador] = lnSeconds
		ENDIF
	ENDFOR

* Output
	IF INLIST(tnDestino,1,2)
		DEBUGOUT lcMensaje
	ENDIF
	IF INLIST(tnDestino,1,3)
		STRTOFILE(lcMensaje + CHR(13) + CHR(10),_Screen.GrabarLog_DevLog,.T.)
	ENDIF
RETURN


DEFINE CLASS MyRun AS Custom
*!*	Values for This.nWindowStyle:
*!*	0 Hides the window and activates another window.
*!*	1 Activates and displays a window. If the window is minimized or maximized, the system restores it to its original size
*!*	  and position. An application should specify this flag when displaying the window for the first time.
*!*	2 Activates the window and displays it as a minimized window. 
*!*	3 Activates the window and displays it as a maximized window. 
*!*	4 Displays a window in its most recent size and position. The active window remains active.
*!*	5 Activates the window and displays it in its current size and position.
*!*	6 Minimizes the specified window and activates the next top-level window in the Z order.
*!*	7 Displays the window as a minimized window. The active window remains active.
*!*	8 Displays the window in its current state. The active window remains active.
*!*	9 Activates and displays the window. If the window is minimized or maximized, the system restores it to its original size
*!*	  and position. An application should specify this flag when restoring a minimized window.
*!*	10 Sets the show-state based on the state of the program that started the application.
	cCommand		= ''
	nWindowStyle	= 9
	lWaitOnReturn	= .F.
	cStdOut			= ''
	cOutputFile	= 'MyGitOutput.txt'	&& File where command output is stored
	oShell			= NULL

	PROCEDURE Init
	LPARAMETERS tcCommand,tnWindowStyle,tlWaitOnReturn
	DECLARE Sleep IN kernel32 INTEGER dwMilliseconds
	* Setting property values:
	* 1. This.Init() receives property values or sets defaults if a parameter is missing
	* 2. Calling program can set the properties after object is instantiated
	* 3. This.Run() receives property values or leaves previous value if a parameter is missing

		* Defaults
		This.cCommand		= IIF(EMPTY(tcCommand),'',tcCommand)
		This.nWindowStyle	= IIF(PARAMETERS() < 2,9,tnWindowStyle)	&& 0 is a valid parameter
		This.lWaitOnReturn	= tlWaitOnReturn
		This.oShell			= CREATEOBJECT("WScript.Shell")
		* If ready to run, do it
		IF PARAMETERS() > 0
	*		This.Run()
		ENDIF

	PROCEDURE Run
	LPARAMETERS tcCommand,tnWindowStyle,tlWaitOnReturn
	* Runs command in command window
		* Defaults
		This.cCommand		= IIF(EMPTY(tcCommand),This.cCommand,tcCommand)
		This.nWindowStyle	= IIF(PARAMETERS() < 2,This.nWindowStyle,tnWindowStyle)		&& 0 is a valid parameter
		This.lWaitOnReturn	= IIF(PARAMETERS() < 3,This.lWaitOnReturn,tlWaitOnReturn)	&& .F. is a valid parameter

		LOCAL loShell AS WScript.Shell,lcCommand,llOk,lnRetVal
		lcCommand = IIF(LEFT(UPPER(This.cCommand),3) # 'CMD','cmd.exe /K ','') + This.cCommand
		llOk	  = .T.
		TRY
			lnRetVal = This.oShell.Run(lcCommand,This.nWindowStyle,This.lWaitOnReturn)
		CATCH
			llOk = .F.
		ENDTRY
		IF !llOk
			MyDebugOut('MyRun.Run - Incorrect parameter. tcCommand = ' + TRANSFORM(tcCommand) ;
							+ ' tnWindowStyle = '+TRANSFORM(tnWindowStyle))
			RETURN -1
		ENDIF

	PROCEDURE RunAppend
	* Runs command, capture STDOUT and ERROUT, append command and output to a file, and show it
	* Following Git convention, it returns 0 to indicate success
	LPARAMETERS tcCommand,tcOutputFile,tlShow
	LOCAL lcTempFile,llOk,lcCommand,loWScript AS WScript,loShell AS WScript.Shell,lnRetVal,lcText1,lcText2,lcText3

		* Defaults

		* Setup
		This.cCommand		= IIF(EMPTY(tcCommand),This.cCommand,tcCommand)
		This.cOutputFile	= IIF(EMPTY(tcOutputFile),This.cOutputFile,tcOutputFile)
		IF TYPE('This.cCommand') # 'C'
			RETURN 
		ENDIF
		lcTempFile = ADDBS(GETENV('Temp')) + SYS(2015) + '.TMP'
		lcTempFile = 'a.txt'
		* Capture STDOUT and ERROUT to lcTempFile
		lcCommand = IIF(LEFT(UPPER(This.cCommand),3) # 'CMD','cmd.exe /C ','') + This.cCommand + [ > "] + lcTempFile + [" 2>&1]

		* Execute
		llOk	= .T.
		TRY
			lnRetVal = This.oShell.Run(lcCommand,0)
		CATCH
			llOk = .F.
		ENDTRY
		IF !llOk
			MyDebugOut('MyRun.RunAppend:  Command failed =  ' + TRANSFORM(lcCommand))
			RETURN -1
		ENDIF

		* Append
		Sleep(1000)	&& Would be better if we wait until program has finished running.  How to do it?
		lcText1 = IIF(FILE(This.cOutputFile),FILETOSTR(This.cOutputFile),'')
		lcText2 = CHR(13) + CHR(10) + '>' + This.cCommand + CHR(13) + CHR(10)
		lcText3 = IIF(FILE(lcTempFile),FILETOSTR(lcTempFile),'')
		* Command redirection inserts LF instead of CR-LF.  Change this for Notepad and other editors.
		IF AT(CHR(13)+CHR(10),lcText3) = 0
			lcText3 = STRTRAN(lcText3,CHR(10),CHR(13)+CHR(10))
		ENDIF			
		llOk = .T.
		TRY
			STRTOFILE(lcText1 + lcText2 + lcText3,This.cOutputFile)
		CATCH
			llOk = .F.
		ENDTRY
		IF !llOk
			? CHR(7)
			WAIT WINDOW 'Could not append, file was in use' TIMEOUT 1.5
			MyDebugOut('MyRun.Run_Append.  Could not append, file was in use')
			RETURN -1
		ENDIF
		Sleep(100)
		TRY
			ERASE (lcTempFile)
			llOk = .T.
		CATCH
			* Don't erase file
		ENDTRY

		* Show
		IF tlShow
			loForm = NEWOBJECT('frmMyGit','mygit.vcx','',This.cOutputFile)
			loForm.Show()
		ENDIF
		RETURN lnRetVal

ENDDEFINE
