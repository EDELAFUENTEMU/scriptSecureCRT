#$language = "VBScript"
#$interface = "1.0"

'#####################################################################
'####### Script automatización procesos en router Teldat   		######
'####### by Eduardo de la Fuente (e.delafuente@outlook.com) 	######
'####### 23-02-2021												######
'#####################################################################


Function AddItem(arr, val) 'Añadir un item a un array redimensionando
    ReDim Preserve arr(UBound(arr) + 1)
    arr(UBound(arr)) = val
    AddItem = arr
End Function

Function VtoStr(arr) 'Convierte un array en str con salto de linea
	Str = ""
	For each linea in arr 
		Str = Str & vbcrlf & linea
    Next
    VtoStr = Str
End Function

Function ListKeyRSA(user, password, hostname, host) 'return un array (resultTxt y resultOperación)
	
	resultTxt = ""
	crt.Screen.WaitForString (user & "@gestib")
	crt.Screen.Send "telnet " & host & chr(13)
	sConnect = crt.Screen.WaitForString("username: ",6) 'Comprobamos si hay conectividad 
		if sConnect = 0  then
			resultOperacion = "	No <-- No hay conexion con el EDC"
			crt.Screen.Send Chr(003)
		else
			crt.Screen.Send user & chr(13)
			crt.Screen.WaitForString "password: "
			crt.Screen.Send password & chr(13)
			sNeumonico = crt.Screen.WaitForString(hostname & " *",6)
			if sNeumonico = 0 then
				resultOperacion = "	No <-- No coincide el hostname del equipo con el hostname insertado previamente"
			else
				crt.Screen.Send "p 5" & chr(13)
				crt.Screen.WaitForString "$"
				crt.Screen.Send "protocol ip" & chr(13)	
				crt.Screen.WaitForString "$"
				crt.Screen.Send "ipsec" & chr(13)
				crt.Screen.WaitForString "$"
				crt.Screen.Send "list key rsa" & chr(13)

				resultTxt = crt.Screen.ReadString("$")
				Set regex = New RegExp
				With regex
					.Global = True
					.Pattern = "(\d rsakey entries)"
				End With 		
				
				Set rsakey = regex.Execute(resultTxt) 
				if rsakey.count > 0 then
					resultOperacion = "	" & rsakey(0).SubMatches(0) & " Yes <-- Obtenido las key"
				else
					resultOperacion = "N/A	Yes <-- Obtenido las key"
				end if
			end if
			
			crt.Screen.Send chr(16) 
			crt.Screen.WaitForString "*"
			crt.Screen.Send  "log" & chr(13)
			crt.Screen.WaitForString chr(13) & "Do you wish to end connection (Yes/No)? "
			crt.Screen.Send "y" & chr(13)
			
		end if
	ListKeyRSA = Array(resultTxt,resultOperacion)
end function

Function FileLists(user, password, hostname, host) 'return un array (resultTxt y resultOperación)
	
	resultTxt = ""
	
	crt.Screen.WaitForString (user & "@gestib")
	crt.Screen.Send "telnet " & host & chr(13)
	sConnect = crt.Screen.WaitForString("username: ",6) 'Comprobamos si hay conectividad 
		if sConnect = 0  then
			resultOperacion = "	No <-- No hay conexion con el EDC"
			crt.Screen.Send Chr(003)
		else
			crt.Screen.Send user & chr(13)
			crt.Screen.WaitForString "password: "
			crt.Screen.Send password & chr(13)
			sNeumonico = crt.Screen.WaitForString(hostname & " *",6)
			if sNeumonico = 0 then
				resultOperacion = "	No <-- No coincide el hostname del equipo con el hostname insertado previamente"
			else
				crt.Screen.Send "p 5" & chr(13)
				crt.Screen.WaitForString "$"
				crt.Screen.Send "file list" & chr(13)	

				resultTxt = crt.Screen.ReadString("Flash Backup")
				resultOperacion = "	Yes <-- Obtenido los datos file"
			end if
			
			crt.Screen.Send chr(16) 
			crt.Screen.WaitForString "*"
			crt.Screen.Send  "log" & chr(13)
			crt.Screen.WaitForString chr(13) & "Do you wish to end connection (Yes/No)? "
			crt.Screen.Send "y" & chr(13)
			
		end if
	FileLists = Array(resultTxt,resultOperacion)
end function

Function ListCertLoadDateExpire(user, password, hostname, host)
	resultTxt = ""
	
	crt.Screen.WaitForString (user & "@gestib")
	crt.Screen.Send "telnet " & host & chr(13)
	sConnect = crt.Screen.WaitForString("username: ",6) 'Comprobamos si hay conectividad 
		if sConnect = 0  then
			resultOperacion = "	No <-- No hay conexion con el EDC"
			crt.Screen.Send Chr(003)
		else
			crt.Screen.Send user & chr(13)
			crt.Screen.WaitForString "password: "
			crt.Screen.Send password & chr(13)
			sNeumonico = crt.Screen.WaitForString(hostname & " *",6)
			if sNeumonico = 0 then
				resultOperacion = "	No <-- No coincide el hostname del equipo con el hostname insertado previamente"
			else
				crt.Screen.Send "p 5" & chr(13)
				crt.Screen.WaitForString "$"
				crt.Screen.Send "protocol ip" & chr(13)	
				crt.Screen.WaitForString "$"
				crt.Screen.Send "ipsec" & chr(13)	
				crt.Screen.WaitForString "$"
				crt.Screen.Send "cert" & chr(13)	
				crt.Screen.WaitForString "$"
				crt.Screen.Send "list loaded-certificates" & chr(13)	
				
				resultTxt = crt.Screen.ReadString("$")
				Set rgx = New RegExp
				With rgx
					.IgnoreCase = True
					.Global     = true
					.Pattern = "(ID[0-9]{2,5}_(T0)?(1|2)\.CER) \(from config\)"
				End With
				
				Set Coincidencias = rgx.Execute(resultTxt) 
				if Coincidencias.count > 0 then
					AuxResultOperacion = ""
										
					For i=0 to Coincidencias.count - 1
					'For Each Coincidencia in Coincidencias
						IdCertificado = Coincidencias(i).SubMatches(0)
											
						crt.Screen.Send "certificate " & IdCertificado & " Print" & chr(13)
						printCert = replace(crt.Screen.ReadString("Publick Key"),chr(13),"")
						resultTxt = resultTxt & printCert
					
						rgx.pattern = "Valid To[ ]*\: (.*) UTC\+"
						expire = ""
						Set expire = rgx.Execute(printCert) 
						AuxResultOperacion = AuxResultOperacion & IdCertificado & " " & expire(0).SubMatches(0) & " | "
					next
						
					ResultOperacion = left(AuxResultOperacion, len(AuxResultOperacion)-2) & "	Yes <-- Datos obtenidos"
				else
					resultOperacion = "	No <-- No se ha encontrado un certificado que cumpla la estructura ID0000-0.CER"
				end if
			end if
			
			crt.Screen.Send chr(16) 
			crt.Screen.WaitForString "*"
			crt.Screen.Send  "log" & chr(13)
			crt.Screen.WaitForString chr(13) & "Do you wish to end connection (Yes/No)? "
			crt.Screen.Send "y" & chr(13)
		end if
	ListCertLoadDateExpire = Array(resultTxt,resultOperacion)
end function

Function GenerateKeyCSR(user, password, hostname, host, ruta)
	resultTxt = ""
	
	Set regex = New RegExp
	With regex
	  .Global = True
	End With 
	
	crt.Screen.WaitForString user & "@gestib"
	crt.Screen.Send "telnet " & host & chr(13)
	sConnect = crt.Screen.WaitForString("username: ",6) 'Comprobamos si hay conectividad 
	if sConnect = 0  then
		resultOperacion = "	No <-- No hay conexion con el EDC"
		crt.Screen.Send Chr(003)
	else
		crt.Screen.Send user & chr(13)
		crt.Screen.WaitForString "password: "
		crt.Screen.Send password & chr(13)
		sNeumonico = crt.Screen.WaitForString(hostname & " *",6)
		if sNeumonico = 0 then
			resultOperacion = "	No <-- No coincide el hostname del equipo con el hostname insertado previamente"
		else
			crt.Screen.Send "p 5" & chr(13)
			crt.Screen.WaitForString "$"
			crt.Screen.Send "protocol ip" & chr(13)
			crt.Screen.WaitForString "$"
			crt.Screen.Send "ipsec" & chr(13)
			crt.Screen.WaitForString "$",8
			crt.Screen.Send "key rsa generate PKI.CER 2048" & chr(13) 'certificado pki de la entidad que firmara el csr
			skey = crt.Screen.ReadString("$") 
			
			regex.Pattern = "It's a good moment.*MAKE ([0-9])\." 
			if regex.test(sKey) = false then
				resultOperacion = "	No <-- La clave no se ha generado correctamente"
			else
				crt.Screen.Send "cert" & chr(13)
				crt.Screen.WaitForString "$"
				crt.Screen.Send "csr" & chr(13)
				crt.Screen.WaitForString "$"
			
				Set nKey = regex.Execute(sKey) 'numero de make X
				n = nKey(0).SubMatches(0)
				crt.Screen.Send "make " & n & chr(13) 'make 1 o make 2, 3, 4
				strkey = crt.Screen.ReadString ("Save in file(Yes/No)?")
				
				regex.Pattern = "-----BEGIN CERTIFICATE REQUEST-----((.|\n|\r|\r\n|\n\r)+)-----END CERTIFICATE REQUEST-----"
				if regex.Test(strkey) <> true then 'Comprueba que el valor recibido cumple el patron ----begin---
					resultOperacion = "	No <-- No se ha generado el CSR"
				else
					crt.Screen.Send "yes" & chr(13) 
					if nKey(0).SubMatches(0) = 1 then
						n = ""
					end if
					crt.Screen.Send "PET" & n & "_CSR.CSR" & chr(13)  
					crt.Screen.WaitForString "$"
					crt.Screen.Send chr(16) 
					crt.Screen.WaitForString "*"
					crt.Screen.Send "p 5" & chr(13) 
					crt.Screen.WaitForString "$"
					crt.Screen.Send "save" & chr(13) 
					crt.Screen.WaitForString "(Yes/No)?"
					crt.Screen.Send "yes" & chr(13)
						
					Set matches = regex.Execute(strkey)
					csr = matches(0).SubMatches(0)
					
					' Quitar espacios en blanco/Sakto línea csr << apaño 
					regex.Pattern = "\r((.|\n)+)\r"
					Set matchestest = regex.Execute(csr)
					csr = matchestest(0).SubMatches(0)
											
					'Genera/sobreescribe Fichero csr
					Set fso = CreateObject("Scripting.FileSystemObject")
					Set csrFile = fso.CreateTextFile( ruta & "\" & hostname & ".csr",True)
					csrFile.WriteLine("-----BEGIN CERTIFICATE REQUEST-----" & chr(13) & chr(10) & csr & "-----END CERTIFICATE REQUEST-----")
					csrFile.Close
					
					resultTxt = "-----BEGIN CERTIFICATE REQUEST-----" & chr(13) & chr(10) & csr & "-----END CERTIFICATE REQUEST-----"
					resultOperacion = "	Yes <--- CSR generado"
				end if
			end if
		end if
		
		crt.Screen.Send chr(16) 
		crt.Screen.WaitForString "*"
		crt.Screen.Send  "log" & chr(13)
		crt.Screen.WaitForString chr(13) & "Do you wish to end connection (Yes/No)? "
		crt.Screen.Send "y" & chr(13)
	end if
	
	GenerateKeyCSR = Array(resultTxt,resultOperacion)
end function

Function UpdateCert(user, password, hostname, host, ruta)
	resultTxt = ""
	pem = ""
	
	Set fso = CreateObject("Scripting.FileSystemObject")
	if (fso.FileExists(ruta & "\" & hostname & ".cer")) <> true then
		resultOperacion = "	No <-- Fichero .Cer no se ha encontrado " 'No existe el fichero .cer
	else
		Set WShell = CreateObject("WScript.Shell") 
		path_openssl = "C:\Program Files\OpenSSL-Win64\bin\openssl.exe"
		strcmd = "cmd /C " & path_openssl & " x509 -inform der -in " & ruta & "\" & hostname & ".cer -outform pem -out " & ruta & "\" & hostname & ".pem "
		WShell.Run(strcmd) 
		crt.sleep 2500
		
		if (fso.FileExists(ruta & "\" & hostname & ".pem")) <> true then
			resultOperacion= "	No <-- No se ha convertido a BASE64 correctamente"
		else
		
			IDCertificado = "ID" & right(left(hostname,7),5) & "_T0" & right(hostname,1) 'W-01111-T01 -> ID01111_01
			Set pemFile = fso.OpenTextFile(ruta & "\" & hostname & ".pem")
			pem = pemFile.ReadAll
			pemFile.Close
			
			crt.Screen.WaitForString user & "@gestib"
			crt.Screen.Send "telnet " & host & chr(13)
			sConnect = crt.Screen.WaitForString("username: ",6) 'Comprobamos si hay conectividad 
			if sConnect = 0  then
				resultOperacion = "	No <-- No hay conexion con el EDC"
				crt.Screen.Send Chr(003)
			else
				crt.Screen.Send user & chr(13)
				crt.Screen.WaitForString "password: "
				crt.Screen.Send password & chr(13)
				sNeumonico = crt.Screen.WaitForString(hostname & " *",6)
				if sNeumonico = 0 then
					resultOperacion = "	No <-- No coincide el hostname del equipo con el hostname insertado previamente"
				else
					crt.Screen.Send "p 5" & chr(13)
					crt.Screen.WaitForString "$"
					crt.Screen.Send "protocol ip" & chr(13)
					crt.Screen.WaitForString "$"
					crt.Screen.Send "ipsec" & chr(13)
					crt.Screen.WaitForString "$"
					crt.Screen.Send "cert" & chr(13)
					crt.Screen.WaitForString "$"
					
					crt.Screen.Send "certificate  " & IDCertificado & ".CER base64" & chr(13) 'base64
					crt.Screen.WaitForString "to escape"
			
					pemLines = Split(pem, vbcrlf)
					For each strLine in pemLines 
						crt.Screen.Send strLine & chr(10)
					next
					crt.Screen.Send chr(13)
					
					resultTxt = ""
					crt.Screen.Send "certificate " & IDCertificado & ".CER load" & chr(13) 'load
					crt.Screen.WaitForString "$"
					crt.Screen.Send "certificate " & IDCertificado & ".CER print" & chr(13)
					resultTxt = resultTxt & replace(crt.Screen.ReadString("Publick Key"),chr(13),"") & vbCrlf
					crt.Screen.Send "list loaded-certificates " & chr(13)
					resultTxt = resultTxt & replace(crt.Screen.ReadString("$"),chr(13),"") & chr(13) & chr(10)
					crt.Screen.Send "exit" & chr(13)
					crt.Screen.WaitForString "$"
					crt.Screen.Send "list key rsa" & chr(13)
					resultTxt =  resultTxt & replace(crt.Screen.ReadString("$"),chr(13),"") & chr(13) & chr(10)
					crt.Screen.Send "exit" & chr(13)
					crt.Screen.WaitForString "$"
					crt.Screen.Send "exit" & chr(13)
					crt.Screen.WaitForString "$"
					crt.Screen.Send "file list" & chr(13)
					resultTxt =  resultTxt & replace(crt.Screen.ReadString("Flash Backup"),chr(13),"")  
					crt.Screen.Send "save" & chr(13)
					crt.Screen.WaitForString "(Yes/No)?"
					crt.Screen.Send "yes" & chr(13) 'yes
					crt.Screen.WaitForString "$"
					
					Set regex = New RegExp
					With regex
						.Global = True
						.Pattern = "Valid To[ ]*\: (.*) UTC\+"
					End With 		
					Set expire = regex.Execute(resultTxt) 
					if expire.count > 0 then
						resultOperacion = "	" & IDCertificado & " Expira " & expire(0).SubMatches(0) & "	Yes <-- Subido correctamente"
					else
						resultOperacion = "	" & IDCertificado & " Expira el N/A	Yes <-- Subido correctamente"
					end if
				end if
				
				crt.Screen.Send chr(16) 
				crt.Screen.WaitForString "*"
				crt.Screen.Send  "log" & chr(13)
				crt.Screen.WaitForString chr(13) & "Do you wish to end connection (Yes/No)? "
				crt.Screen.Send "y" & chr(13)
			end if
		end if
	end if
	
	UpdateCert = Array(resultTxt,resultOperacion)
end function

sub main()
	user = "<mi_usuario>"
	password = "<mi_contraseña>"
	hostBastion = "<mi_ip_host_bastion>" 
	
	Set fso = CreateObject("Scripting.FileSystemObject")
	Set regex = New RegExp
	With regex
		.IgnoreCase = True
		.Global     = False
		.Pattern = "^(W-[0-9]{5}-T0(1|2)) ([0-9]{1,3}.[0-9]{1,3}.[0-9]{1,3}.[0-9]{1,3})$" 'regex del hostname e IP Host. Ej: W-01111-T02 192.0.0.1
	End With

	'Creamos la conexion
	crt.Screen.Synchronous = True
	cmd =  "/SSH2 /L " & user & " /PASSWORD " & password &  " " & hostBastion 
	crt.Session.Connect cmd 'Dara error si no consigue establecer conexión con GestIberc
	
	'Pregunta la carpeta de trabajo
	Set ShellApp = CreateObject("Shell.Application")
	Set rUnidad = ShellApp.BrowseForFolder(0, "Seleccione la carpeta de trabajo", 16384)
	ruta = rUnidad.Self.Path 
	
	'Array input y output
	vOutput = array() 'Contendra un array multidimensional con la estrucutura {1:[resultTxt y resultOperación],2:[resultTxt y resultOperación]}
	vInput1 = array() 'contendra los equipos afectados
	
	if fso.FileExists(ruta & "\lote.txt") and (MsgBox ("Detectado fichero trabajo por lote.¿Desea usarlo?", vbYesNo + vbQuestion, "Confirmacion lote" ) = vbYes) then
		Set job = fso.OpenTextFile(ruta & "\lote.txt",1) 
		Do While job.AtEndOfStream <> True
			vInput1 = AddItem(vInput1, job.ReadLine)
		loop
	else
		vInput1 = AddItem(vInput1, InputBox("No se ha detectado fichero para procesar por lotes. Indique el hostname e IP host siguiendo la siguiente estructura" & vbCrlf & "W-01111-T01 192.168.0.1" ,"EDC afectado"))
	end if
	
	'Pregunta que tipo de operación quieres hacer
	vTipoOperacion = array(": File List",": List Key RSA",": Certificados Caducidad",": Generar CSR",": Cargar CER")
	sTipo = "Que quiere ejecutar (indique numero) " & vbCrlf
	for i=0 to ubound(vTipoOperacion)
		sTipo = sTipo & i + 1 & vTipoOperacion(i) & vbCrlf
	next
	sTipo = InputBox(sTipo,"Tipo de Operación")

	
	'Confirmación operacion y equipos afectados
	Str = "Operacion seleccionada: " & vTipoOperacion(sTipo-1) & vbCrlf & "EDC afectados: " & vbCrlf
	For each linea in vInput1 
		Str = Str & vbcrlf & linea
    Next
	if MsgBox (Str, vbYesNo + vbQuestion, "Confirmacion") = vbNo then
		MsgBox ( "Operacion suspendida!!!" )
	else
		'Inicio de la operacion
		for each linea in vInput1 
			vAux = array("","	No <-- La línea no cumple la estructura predifinida") 'Valor por defecto
			if	regex.Test(linea) then 'si la línea cumple la estrucuta esencial
				vInput = Split(linea, " ")
				hostname = vInput(0)
				host = vInput(1)
				
				select Case sTipo
					case 1:
						vAux = FileLists(user, password, hostname, host) 'return un array (resultTxt y resultOperación)
					case 2:
						vAux = ListKeyRSA(user, password, hostname, host) 'return un array (resultTxt y resultOperación)
					case 3:
						vAux = ListCertLoadDateExpire(user, password, hostname, host) 'return un array (resultTxt y resultOperación)
					case 4:
						vAux = GenerateKeyCSR(user, password, hostname, host, ruta) 'return un array (resultTxt y resultOperación) y crea un fichero petN_csr.csr
					case 5:
						vAux = UpdateCert(user, password, hostname, host, ruta) 'return un array (resultTxt y resultOperación) y genera un fichero .pem
				end select
				
				vAux(1) = hostname & "	" & host & "	" & vAux(1) 'Incluye hostname e ipgestion en el resultado de la operación
			else
				vAux(1) = linea & "	" & vAux(1) 'Incluye la línea mal formateada en el resultado de la operación
			end if
		
			vOutput = AddItem(vOutput, vAux) 'Linea con el resultado de la operación
		next
		crt.Screen.Send "Tarea Terminada" 
		
		'Genera fichero con los resultado con los valores devuelto en output
		crlf = chr(13) & chr(10)
		Set FileOutput = fso.CreateTextFile( ruta & "\resultado.txt",True)
		FileOutput.WriteLine("##################################################" & crlf & Date & " Resultado abreviado (Operación tipo" & vTipoOperacion(sTipo-1) & ")" & crlf ) 
		for x=0 to UBound(vOutput)   
			FileOutput.WriteLine(vOutput(x)(1)) 
		Next
		FileOutput.WriteLine(crlf & "##################################################" & crlf & Date & " Resultado Ampliado (Operación tipo" & vTipoOperacion(sTipo-1) & ")" & crlf) 
		for x=0 to UBound(vOutput)   
			FileOutput.WriteLine(vOutput(x)(1) & crlf)
			FileOutput.WriteLine(vOutput(x)(0) & crlf & "**************************************************" & crlf & crlf)
		Next
		FileOutput.Close
	end if
end sub