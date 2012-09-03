'***********************************************************************************************************************
' Arquivo: script.vbs
' Autor: Henrique Grammelsbacher
' Data:   05-feb-08
' Ult at: 29/6/2012 17:04:43
' Description: Pacote de atualizacao generico do SDM

Option Explicit
'***********************************************************************************************************************
'Define as variaveis de execucao do script de atualizacao do Service Desk
Dim fso, ows, ocmds, oNxEnv, ostd
Dim strNxroot, strSite, strPath, strLogPath, msg, iTimeOut, strProvider, sLinha,strWinVersion
Set fso = CreateObject("Scripting.FileSystemObject")
Set ows = CreateObject("WScript.Shell")
set ostd = WScript.StdOut

'Dependendo do SO, a chave do SDM muda, portanto, vamos na tentativa e erro
If regExists("HKEY_LOCAL_MACHINE\SOFTWARE\ComputerAssociates\CA Service Desk\Install Path") = True Then 
    strNxroot = ows.RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\ComputerAssociates\CA Service Desk\Install Path")
    
Elseif regExists("HKEY_LOCAL_MACHINE\SOFTWARE\CA Service DESK") = True Then 
    strNxroot = ows.RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\CA Service DESK\Install Path\Install Path")
    
Elseif regExists("HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\ComputerAssociates\CA Service Desk\Install Path") = True Then 
    strNxroot = ows.RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\ComputerAssociates\CA Service Desk\Install Path")
    
Else
    msgbox "Não foi possível encontra a chave de registro do CA Service Desk. Entre em contato com a Interadapt.", vbOK, "Erro"
    wscript.quit 
    
End If 

if right(strNxroot,1) = "\" then
	strNxroot = mid(strNxroot, 1, len(strNxroot)-1)
end if
strSite = strNxroot & "\site"
strPath = replace(WScript.ScriptFullName, WScript.ScriptName, "")
strPath = mid(strPath, 1, len(strPath) - 1)
strLogPath = strPath & "\logs"  
iTimeOut = 120
Set ocmds = fso.CreateTextFile(strPath & "\logs\cmds.log", True)

'Determina o provider de banco... Por causa do Marcos e do Black, tive que iterar pelo NX.env
strProvider = ""
If fso.FileExists(strNxroot & "\NX.env") Then
    Set oNxEnv = fso.OpenTextFile(strNxroot & "\NX.env")
    Do While Not oNxEnv.AtEndOfStream
        sLinha = oNxEnv.readLine
        if len(sLinha) > 22 then
            if left(sLinha, 15) = "@NX_JDBC_DRIVER" then
                if instr(ucase(sLinha), "ORACLE") > 0 then
                    strProvider = "Oracle"
                elseif instr(ucase(sLinha), "SQLSERVER") > 0 then
                    strProvider = "SQL Server"
                else
                    strPrivider = ""
                end if
                exit Do
            end if

        end if

    Loop
    if strProvider = "" then
        msgbox "Não foi possível determinar o provider de banco. Entre em contato com a Interadapt", vbOK, "Erro"
        wscript.quit 
    end if
else
    msgbox "Não foi possível encontrar o arquivo NX.env no caminho:" & vbcrlf &  strNxroot & "\NX.env", vbOK, "Erro"
    wscript.quit 
End If

verficaAplicacao()


'Para o service desk
runCmd "pdm_halt"

'Verifica se o servico parou corretamente, se nao consegui, aborta.
If verificaPdmDown() Then

    'bkp dos arquivos
    bkp

    'Exclui arquivos desnecessários e Atualiza arquivos do Service Desk
    deleteFiles
    copyFiles

    'Inicializa em modo de manutencao
    runCmd "pdm_d_mgr -s DBADMIN"
    
    'Extrai as tabelas que futuramente serão carregadas
    extractFiles

    'Backup especifico das tabelas a serem modificadas
    pdm_bkp_userloads   
    
    'Caso exista o arquivo de load wspcol.userload, faz atualização do modelo de objeto
    if fso.fileExists(strPath & "\Userload\wspcol.userload") then
		If verificaPdmDBAdmin(strProvider) Then

	        'Faz o backup FULL do banco
	        'pdm_backup
	        
	        'Caso existam tabelas que serão reconstruidas, efetue o extract aqui
	        'pdm_extract "Events"
	        'pdm_extract "Spell_Macro"

	        'Load das modificações do modelo de objeto
	        pdm_userload "wspcol.userload"
	        pdm_userload "wsptbl.userload"

	        'Para o servico para merge dos arquivos
	        runCmd "pdm_halt"
	        
	        If verificaPdmDown() Then
	            
	            'Faz a publicação do schema
	            runCmd "pdm_publish"
	            
	        End If
	    
	    End If
	    
	    'O build das tabelas a serem reconstruidas manualmente deve ser feito aqui
	    'build "Events"
	    'build "Spell_Macro"

	    'Inicializa em modo de manutencao
	    runCmd "pdm_d_mgr -s DBADMIN"
	end if

    'Verifica se o modo de manutencao subiu corretamente
    If verificaPdmDBAdmin(strProvider) Then

        'Recarrega as tabelas que foram reconstruidas manualmente
        'pdm_restore "Events"
        'pdm_restore "Spell_Macro"
            		
		'Carga dos dados
		carregar_tudo
	
        'Re-Instala as options
        runCmd "pdm_options_mgr"
	
    End If
        
    runCmd "pdm_halt"
        
    'Se terminou a mautencao, re-inicializa o servico
    If verificaPdmDown() Then
        runCmd "net start pdm_daemon_manager"
    End If
    
    'Verifica se o USvD subiu corretamente e exibemensgem de sucesso ou fracasso
    If verificaPdmUp() Then
        msgbox "Sistema atualizado normalmente." & vbcrlf & "Por favor, envie o conteudo da pasta logs para analise", 1, "Operacao efetuada com sucesso"        
    End If
    
End If

oNxEnv.close

Set ows = nothing
Set fso = nothing
set oNxEnv = nothing
set ostd = nothing




'***********************************************************************************************************************
'Funcoes acessorias
'***********************************************************************************************************************
'***********************************************************************************************************************
' Funcao: bkp()
' Autor: Henrique Grammelsbacher
' Data: 09-04-08
' Description: Efetua o bakup das pastas enumeradas em bkp.txt
Function pdm_backup()

    Dim cmd, hora
    hora = retornaData
    runCmdwnlog "pdm_extract ALL > """ & strPath & "\bkp\full.bkp"" 2> " & strPath &  "\logs\" & hora & "bkpfull.log"
    
End Function

'***********************************************************************************************************************
' Funcao: build()
' Autor: Henrique Grammelsbacher
' Data: 09-04-08
' Description: Executa o sqlbuild ou orclbuild, conforme o banco em uso
Function build(tabela)

    Dim cmd, hora, pref
    hora = retornaData
    if strProvider = "Oracle" then
        pref = "orcl"
        else
        pref = "sql"
    end if
    runCmdwnlog pref & "build -C -p " & tabela & " mdb """ & strSite & "\ddict.sch"" < y.txt > " & strPath &  "\logs\" & hora & "build.log 2> " & strPath &  "\logs\" & hora & "build.err"
    
End Function


'***********************************************************************************************************************
' Funcao: bkp()
' Autor: Henrique Grammelsbacher
' Data: 09-04-08
' Description: Efetua o bakup das pastas enumeradas em bkp.txt
Function bkp()
    
    Dim cmd, sRes, sArq
    Dim oFile, oLogFile
    
    Set oLogFile = fso.CreateTextFile(strPath & "\logs\bkp.log", True)
    oLogFile.writeLine "Inicializando o bkp de arquivos - " & now()
    
    
    If fso.FileExists(strPath & "\bkp.txt") Then
        Set oFile = fso.OpenTextFile(strPath & "\bkp.txt")
        Do While Not oFile.AtEndOfStream
            sArq = oFile.readLine
            If fso.FileExists(strNxroot & "\" & sArq) Or fso.FolderExists(strNxroot& "\" & sArq) Then
                cmd = strPath & "\uteis\rar.exe a -ag-YYYY-MM-DD """ & strPath & "\bkp\arq.rar""" & " """ & strNxroot & "\" & sArq & Chr(34)
                runCmdwLog cmd, "arq.rar"
                oLogFile.writeline "-" & now() & " - " & cmd
            Else
                oLogFile.writeline "-" & now() & " - File not Found - " & strNxroot & "\" & sArq
            End If
            
        Loop
    End If
    
    oLogFile.writeLine "Fim do bkp de arquivos - " & now()
    
    oLogFile.Close
    Set oLogFile = nothing
    
End Function





'***********************************************************************************************************************
' Funcao: pdm_deref()
' Autor: Henrique Grammelsbacher
' Data: 05-fev-2008
' Description: Efetua o userload dos arquivos dat da pasta userload
Function pdm_deref(script, userload, dat)
    Dim cmd, hora
    hora = retornaData
    cmd = "pdm_deref -p -s " & strPath & "\deref\" & script & " " & strPath & "\userload\" & userload & " > " & strPath & "\userload\" & dat & " 2> " & strPath & "\logs\" & hora & "deref_" & userload & ".err"
    runCmdwnLog cmd
End Function

'***********************************************************************************************************************
' Funcao: sed()
' Autor: Henrique Grammelsbacher
' Data: 15-abr-2008
' Description: Executa um script sed na pasta sed/
Function sed(script, dat, dest)
    Dim cmd, hora
    hora = retornaData
    cmd = strPath & "\uteis\sed -f " & strPath & "\sed\" & script & " " & strPath & "\userload\" & dat & " > " & strPath & "\userload\" & dest & " 2> " & strPath & "\logs\" & hora & "sed_" & dest & ".err"
    runCmdwnLog cmd
End Function


'***********************************************************************************************************************
' Funcao: pdm_userload()
' Autor: Henrique Grammelsbacher
' Data: 05-fev-2008
' Description: Efetua o userload dos arquivos dat da pasta userload
Function pdm_userload(arquivo)
    Dim cmd
    cmd = "pdm_userload -v -f " & strPath & "\userload\" & arquivo
    runCmdwLog cmd, "pdm_userload_" & arquivo
End Function

'***********************************************************************************************************************
' Funcao: pdm_delete()
' Autor: Henrique Grammelsbacher
' Data: 9/4/2010 12:08:31 PM
' Description: Efetua o userload -r dos arquivos dat da pasta userload
Function pdm_delete(arquivo)
    Dim cmd
    cmd = "pdm_userload -v -r -f " & strPath & "\userload\" & arquivo
    runCmdwLog cmd, "pdm_userload_" & arquivo
End Function


'***********************************************************************************************************************
' Funcao: pdm_restore()
' Autor: Henrique Grammelsbacher
' Data: 05-fev-2008
' Description: Efetua o userload de uma tabela backupeada com function pdm_backup
Function pdm_restore(arquivo)
    Dim cmd
    cmd = "pdm_userload -v -f " & strPath & "\bkp\" & arquivo & ".bkp"
    runCmdwLog cmd, "pdm_userload_" & arquivo
End Function


'***********************************************************************************************************************
' Funcao: pdm_extract()
' Autor: Henrique Grammelsbacher
' Data: 09-abr-2008
' Description: Efetua o extract de uma tabela
Function pdm_extract_toload(tabela, arquivo)
    Dim cmd
    Dim hora
    hora = retornaData
    cmd = "pdm_extract -f""" & tabela & """ > " & strPath & "\userload\" & arquivo  & " 2> " & strPath &  "\logs\" & hora & "_" & arquivo & ".log"
    runCmdwnLog cmd
End Function

Function pdm_extract(tabela)
    Dim cmd
    Dim hora
    hora = retornaData
    cmd = "pdm_extract " & tabela & " > " & strPath & "\bkp\" & tabela & ".bkp 2> " & strPath &  "\logs\" & hora & "_" & tabela & ".log"
    runCmdwnLog cmd
End Function

'***********************************************************************************************************************
' Funcao: carregar()
' Autor: Henrique Grammelsbacher
' Data: 2:10 PM 8/26/2008
' Description: Identifica, através do nome do arquivo, o que fazer com ele
Function carregar_tudo()

	dim ofiles, ofile

	set ofiles = fso.getFolder(strPath & "\UserLoad")
	for each ofile in ofiles.files

		if ofile.name <> "wspcol.userload" and ofile.name <> "wsptbl.userload" and ofile.name <> "info.txt" and ofile.name <> "" then
		
			carregar ofile.name
		
		end if
	next 

	
	set ofiles = nothing
	set ofile = nothing

End Function

'***********************************************************************************************************************
' Funcao: carregar()
' Autor: Henrique Grammelsbacher
' Data: 05-fev-2008
' Description: Identifica, através do nome do arquivo, o que fazer com ele
Function carregar(arq)

	dim strDeref, strArqC
	if len(arq) < 4 then
	   carregar = 0
	   exit function
	end if

	strDeref = left(arq, len(arq) - 3) & "deref"

	if mid(arq, len(arq) - 5, 2)  = "_L" then

		strArqC = left(arq, len(arq) - 3) & "load"
		pdm_deref strDeref, arq, strArqC
		pdm_load strArqC

	elseif mid(arq, len(arq) - 5, 2)  = "_U" then

		strArqC = left(arq, len(arq) - 3) & "userload"
		pdm_deref strDeref, arq, strArqC
		pdm_userload strArqC

	elseif mid(arq, len(arq) - 5, 2)  = "_R" then

		strArqC = left(arq, len(arq) - 3) & "replace"
		pdm_deref strDeref, arq, strArqC
		pdm_replace strArqC

	elseif mid(arq, len(arq) - 5, 1)  = "S" then

        if mid(arq, len(arq) - 4, 1)  = "L" then
    		strArqC = left(arq, len(arq) - 3) & "tmp"
    		pdm_deref strDeref, arq, strArqC
            sed left(arq, len(arq) - 3) & "sed", strArqC, left(arq, len(arq) - 3) & "load" 
    		pdm_load left(arq, len(arq) - 3) & "load"
        elseif mid(arq, len(arq) - 4, 1)  = "R" then
    		strArqC = left(arq, len(arq) - 3) & "tmp"
    		pdm_deref strDeref, arq, strArqC
            sed left(arq, len(arq) - 3) & "sed", strArqC, left(arq, len(arq) - 3) & "replace" 
    		pdm_replace left(arq, len(arq) - 3) & "replace"
        elseif mid(arq, len(arq) - 4, 1)  = "D" then
    		strArqC = left(arq, len(arq) - 3) & "tmp"
    		pdm_deref strDeref, arq, strArqC
            sed left(arq, len(arq) - 3) & "sed",  strArqC, left(arq, len(arq) - 3) & "delete" 
    		pdm_delete left(arq, len(arq) - 3) & "delete"
        else
    		strArqC = left(arq, len(arq) - 3) & "tmp"
    		pdm_deref strDeref, arq, strArqC
            sed left(arq, len(arq) - 3) & "sed",  strArqC, left(arq, len(arq) - 3) & "userload" 
    		pdm_userload left(arq, len(arq) - 3) & "userload"
        end if

	elseif mid(arq, len(arq) - 5, 1)  = "D" then

        if mid(arq, len(arq) - 4, 1)  = "L" then
    		strArqC = left(arq, len(arq) - 3) & "tmp"
            sed left(arq, len(arq) - 3) & "sed",  arq, strArqC 
    		pdm_deref strDeref, strArqC, left(arq, len(arq) - 3) & "load"
    		pdm_load left(arq, len(arq) - 3) & "load"
        elseif mid(arq, len(arq) - 4, 1)  = "R" then
    		strArqC = left(arq, len(arq) - 3) & "tmp"
            sed left(arq, len(arq) - 3) & "sed",  arq, strArqC 
    		pdm_deref strDeref, strArqC, left(arq, len(arq) - 3) & "replace"
    		pdm_replace left(arq, len(arq) - 3) & "replace"
        elseif mid(arq, len(arq) - 4, 1)  = "D" then
    		strArqC = left(arq, len(arq) - 3) & "tmp"
            sed left(arq, len(arq) - 3) & "sed",  arq, strArqC 
    		pdm_deref strDeref, strArqC, left(arq, len(arq) - 3) & "delete" 
    		pdm_delete left(arq, len(arq) - 3) & "delete"
        else
    		strArqC = left(arq, len(arq) - 3) & "tmp"
            sed left(arq, len(arq) - 3) & "sed",  arq, strArqC 
    		pdm_deref strDeref, strArqC, left(arq, len(arq) - 3) & "userload" 
    		pdm_userload left(arq, len(arq) - 3) & "userload"
        end if
		
	else
		
		if ucase(right(arq, 8)) = "USERLOAD" then
		
			pdm_userload arq
		
		elseif ucase(right(arq, 7)) = "REPLACE" then
		
		    pdm_replace arq
		
		elseif ucase(right(arq, 6)) = "DELETE" then
		
		    pdm_delete arq
		
		elseif ucase(right(arq, 4)) = "LOAD" then
		
			pdm_load arq
		
		else
		
		  oLogFile.writeLine "O arquivo '" & arq & "' nao foi carregado pois nao possui uma extensao valida."
		
		end if

	end if

End Function

'***********************************************************************************************************************
' Funcao: pdm_load()
' Autor: Henrique Grammelsbacher
' Data: 05-fev-2008
' Description: Efetua o userload dos arquivos dat da pasta userload
Function pdm_load(arquivo)
    Dim cmd
    cmd = "pdm_load -i -v -f " & strPath & "\userload\" & arquivo
    runCmdwLog cmd, "pdm_load_" & arquivo
End Function

'***********************************************************************************************************************
' Funcao: pdm_replace()
' Autor: Henrique Grammelsbacher
' Data: 05-fev-2008
' Description: Efetua o userload dos arquivos dat da pasta userload
Function pdm_replace(arquivo)
    Dim cmd
    cmd = "pdm_replace -v -f " & strPath & "\userload\" & arquivo
    runCmdwLog cmd, "pdm_replace_" & arquivo
End Function

'***********************************************************************************************************************
' Funcao: extractFiles()
' Autor: Henrique Grammelsbacher
' Data: 3-mai-2009 6:25:12 PM
' Description: Extrai arquivos para serem carregados
Function extractFiles()
    
    Dim cmd, sRes, sArq
    Dim oFile, oLogFile, sSelect, sNewFile, aTexto
    
    Set oLogFile = fso.CreateTextFile(strPath & "\logs\extracts.log", True)
    oLogFile.writeLine "Inicializando a export de tabelas - " & now()
    
    
    If fso.FileExists(strPath & "\extractToLoad.txt") Then
        Set oFile = fso.OpenTextFile(strPath & "\extractToLoad.txt")
        Do While Not oFile.AtEndOfStream
            sArq = oFile.readLine
            aTexto = split(sArq, "||")
            if ubound(aTexto) = 1 then
                sSelect = trim(aTexto(0))
                sNewFile = trim(aTexto(1)) 
                pdm_extract_toload sSelect, sNewFile
            end if           
        Loop
    End If
    
    oLogFile.writeLine "Fim da exclusao de arquivos - " & now()
    
    oLogFile.Close
    Set oLogFile = nothing
    
End Function


'***********************************************************************************************************************
' Funcao: deleteFiles()
' Autor: Henrique Grammelsbacher
' Data: 05-fev-2008
' Description: Exclui os arquivos enumerados em deletar.txt
Function deleteFiles()
    
    Dim cmd, sRes, sArq
    Dim oFile, oLogFile
    
    Set oLogFile = fso.CreateTextFile(strPath & "\logs\deleteFiles.log", True)
    oLogFile.writeLine "Inicializando a exclusao de arquivos - " & now()
    
    
    If fso.FileExists(strPath & "\deletar.txt") Then
        Set oFile = fso.OpenTextFile(strPath & "\deletar.txt")
        Do While Not oFile.AtEndOfStream
            sArq = oFile.readLine
            If fso.FileExists(strNxroot & sArq) Then
                sRes = fso.DeleteFile(strNxroot & sArq)
                oLogFile.writeline "-" & now() & " - " & sRes & " - " & sArq
            Else
                oLogFile.writeline "-" & now() & " - File not Found - " & sArq
            End If
            
        Loop
    End If
    
    oLogFile.writeLine "Fim da exclusao de arquivos - " & now()
    
    oLogFile.Close
    Set oLogFile = nothing
    
End Function

'***********************************************************************************************************************
' Funcao: copyFiles()
' Autor: Henrique Grammelsbacher
' Data: 05-fev-2008
' Description: Copia os arquivos de nxroot para a pasta do Service Desk
Function copyFiles()
    Dim cmd
    
    cmd = strPath & "\uteis\robocopy /s  " & """" & strPath & "\nxroot" & """" & " " & """" & strNxroot & """"
    runCmdwLog cmd, "robocopy"

End Function

'***********************************************************************************************************************
' Funcao: runCmd()
' Autor: Henrique Grammelsbacher
' Data: 05-fev-2008
' Description: Executa comando e cria log do output
Function runCmd(cmd)
    runCmdwLog cmd, cmd
End Function

Function runCmdwnLog(cmd)
    Dim msg, cmd1
    cmd1 = "cmd /c " & cmd
    ocmds.writeline "EXEC - " & now & " - " & msg & " - " & cmd
    WScript.StdOut.WriteLine(">> " & cmd) 
    msg = ows.run (cmd1, 0, true) 
    If ( msg > 0 and instr(cmd, "robocopy") <= 0 )  OR ( msg > 4 and instr(cmd, "robocopy") > 0 ) Then
        If Not continua(cmd) Then
            wscript.quit
        End If 
    End If 
End Function

Function runCmdwLog(cmd, logfile)
    Dim re, hora, errLog, exeLog, owss, cmd1

    Set re = New RegExp
    re.pattern = " "
    re.global = true
    
    hora = retornaData
    errLog = """" & strPath & "\logs\" & hora & re.replace(logfile, "_") & ".err" & """"
    exeLog = """" & strPath & "\logs\" & hora & re.replace(logfile, "_") & ".log" & """"
    cmd1 = "cmd /c " & cmd &  " > " & exeLog & " 2> " & errLog
    WScript.StdOut.WriteLine(">> " & cmd) 
    ocmds.writeline "EXEC - " & now & " - " & msg & " - " & cmd
    msg = ows.run (cmd1 ,0, true)
    
    Set re = nothing

    If ( msg > 0 and instr(cmd, "robocopy") <= 0 )  OR ( msg > 4 and instr(cmd, "robocopy") > 0 ) Then
        If Not continua(cmd) Then
            wscript.quit
        End If 
    End If 
    
End Function

'***********************************************************************************************************************
' Funcao: verificaPdmDown()
' Autor: Henrique Grammelsbacher
' Data: 04-fev-2008
' Description: Aguarda 60 seg pela parada do Service Desk, ou seja, o sslump foi eliminado da task list
Function verificaPdmDown()
    If aguardaHalt(iTimeOut) Then 
        verificaPdmDown = true
    Else
        msgbox "Não foi possível executar a manutenção. Por favor, reinicie o serviço 'Unicenter Service Desk' manualmente" & vbcrlf & "e envie a pasta logs para avaliacao", 16, "Erro fatal"
        wscript.quit
    End If 
End Function

'***********************************************************************************************************************
' Funcao: verificaPdmUp()
' Autor: Henrique Grammelsbacher
' Data: 04-fev-2008
' Description: Aguarda o tempo em segs definido em iTimeOut pela inicialização do tomcat do Service Desk, ou seja, o pdm_tomcat_nxd foi carregado
Function verificaPdmUp()
    If aguardaStart(iTimeOut, "pdm_tomcat_nxd") Then 
        verificaPdmUp = true
        Else
        msgbox "Não foi possível executar a manutenção. Por favor, reinicie o serviço 'Unicenter Service Desk' manualmente" & vbcrlf & "e envie a pasta logs para avaliacao", 16, "Erro fatal"
        wscript.quit
    End If 
End Function

'***********************************************************************************************************************
' Funcao: verificaPdmDBAdmin()
' Autor: Henrique Grammelsbacher
' Data: 04-fev-2008
' Description: Aguarda em segs definido em  iTimeOut seg pela inicializacao do service desk em DB Admin Mode
Function verificaPdmDBAdmin(prov)
    
	dim strProc
	if prov = "Oracle" then
	
		strProc = "orcl_prov_nxd"
		
		else
		
		strProc= "sql_prov_nxd"
	
	end if
	
	If aguardaStart(iTimeOut, strProc) Then 
        verificaPdmDBAdmin = true
        Else
        msgbox "Não foi possível iniciar o serviço do Service Desk - Entre em contato com a Interadapt", 16
        wscript.quit
    End If 
End Function

'***********************************************************************************************************************
' Funcao: aguardaHalt()
' Autor: Henrique Grammelsbacher
' Data: 04-fev-2008
' Description: Aguarda xx segundos pelo processo slump na lista de tarefas (task manager)
Function aguardaHalt(segundos) 
    Dim objWMIService, objProcess, colProcess
    Dim strComputer, strList, intLoop
    
    strComputer = "."
    
    Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" _
    & strComputer & "\root\cimv2")
    
    strList = "Inicio" 
    intLoop = 0
    
    Do While len(strList) > 5 And intLoop < segundos
        
        strList = ""
        Set colProcess = objWMIService.ExecQuery ("Select * from Win32_Process")
        
        For Each objProcess In colProcess
            If instr(objProcess.Name, "slump") > 0 Then
                strList = strList & objProcess.Name
            End If
        Next
        
        wscript.sleep 1000
        intLoop = intLoop + 1
        
    Loop

    If intLoop >= segundos Then
        aguardaHalt = False
        Else
        aguardaHalt = true 
    End If
    
    Set objWMIService = nothing
    
End Function
    
    
'***********************************************************************************************************************
' Funcao: aguardaHalt()
' Autor: Henrique Grammelsbacher
' Data: 04-fev-2008
' Description: Aguarda xx segundos para o processo aparecer na lista de tarefas (task manager)
Function aguardaStart(segundos, processo) 
    Dim objWMIService, objProcess, colProcess
    Dim strComputer, strList, intLoop
	
    strComputer = "."
    
    Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" _
    & strComputer & "\root\cimv2")
    
    strList = "" 
    intLoop = 0
    
    Do While len(strList) = 0 And intLoop < segundos
        
        Set colProcess = objWMIService.ExecQuery ("Select * from Win32_Process")
        
        For Each objProcess In colProcess
            If instr(objProcess.Name, processo) > 0 Then
                strList = strList & objProcess.Name
            End If
        Next
        
        wscript.sleep 1000
        intLoop = intLoop + 1
        
    Loop

    If intLoop >= segundos Then
        aguardaStart = False
        Else
        aguardaStart = true 
    End If
    
    Set objWMIService = nothing
    
End Function    

'***********************************************************************************************************************
' Funcao: verificaServicoPdm()
' Autor: Henrique Grammelsbacher
' Data: 04-fev-2008
' Description: Consulta o service manager pelo status de um servico
Function verificaServicoPdm()

    'Status dos servicos
    '1 = stopped
    '2 = start pending
    '3 = stop pending
    '4 = running
    '5 = continue_poending
    '6 = pause_pending
    '7 = paused
    '8 = error

    Dim objComputer, aService

    Set objComputer = GetObject("WinNT://.,computer")
    objComputer.Filter = Array("Service")

    verificaServicoPdm = 8
    For Each aService In objComputer
        If instr(aService.Name, "pdm_daemon_manager") Then
           On Error Resume Next
           verificaServicoPdm =  aService.Status
        End If
    Next
    
    Set objComputer = nothing
    


End Function    

'***********************************************************************************************************************
' Funcao: continua()
' Autor: Henrique Grammelsbacher
' Data: 09-abr-2008
' Description: Pergunta ao usuario se quer continuar depois de um erro

Function continua(msg)
    Dim res
    
    res = msgbox("O comando abaixo foi executado e retornou um código de erro. Deseja continuar?" & vbcrlf & msg, vbOKCancel, "Erro")    
    If res = 1 Then
        continua = true
        Else
        continua = False
    End If
    
End Function

function retornaData()

	retornaData = formata(year(now())) & "-" & formata(month(now())) & "-" & formata(day(now())) & "--" & formata(hour(now())) & "-" & formata(minute(now())) & "-" & formata(second(now())) & "_"

end function
	

function formata(num)

	dim strNum

	if len(num) = 1 then
		strNum = "0"  & num
		else
		strNum = num
	end if

	formata = strNum

end function



'***********************************************************************************************************************
' Funcao: pdm_bkp_userloads()
' Autor: Henrique Grammelsbacher
' Data: 09-abr-2008
' Description: Pergunta ao usuario se quer continuar depois de um erro
Function pdm_bkp_userloads()
    Dim cmd
    Dim hora
    dim oFile, oLogFile
    dim sArq

    Set oLogFile = fso.CreateTextFile(strPath & "\logs\bkp.log", True)
    oLogFile.writeLine "Inicializando o bkp de arquivos - " & now()

    
    runCmdWnLog """" & strPath & "\uteis\criaListaTabsUserlod.bat""" 
    If fso.FileExists(strPath & "\bkp\lst_bkp_tab.txt") Then
        Set oFile = fso.OpenTextFile(strPath & "\bkp\lst_bkp_tab.txt")
        Do While Not oFile.AtEndOfStream
            sArq = trim(oFile.readLine)
            hora = retornaData
            cmd = "pdm_extract " & sArq & " > " & strPath & "\bkp\" & sArq  & ".bkp 2> " & strPath &  "\logs\" & hora & "_" & sArq & ".log"
            runCmdwnLog cmd
            oLogFile.writeline "-" & now() & " - " & cmd
        Loop
    End If
        
    oFile.close
    oLogFile.close

    
    set oFile = nothing
    set oLogFile = nothing

End Function


'***********************************************************************************************************************
' Funcao: regExists()
' Autor: Marcos Strapazon
' Data: 11-maio-2010
' Description: Verifica se uma chave de registro existe
function regExists (regKey)
	Dim WshShell, Root
	Set WshShell = CreateObject("WScript.Shell")
	On Error Resume Next
	regExists = WSHShell.RegRead (regKey)
	if not isEmpty(regExists) then 
        regExists=true
    else
        regExists=false
    end if
    err.clear
	on error Goto 0
	Set WSHShell = Nothing
end function
 
'***********************************************************************************************************************
' Funcao: verficaAplicacao ()
' Autor: Lucas Guimaraes
' Data: 29-jun-2012
' Description: Verifica se o pacote jah foi aplicado
function verficaAplicacao ()
    Dim oPkgFile
    Dim pkgName, ctrlFilePath
    Dim control
    Dim FileContents, LinePart, LineParts
    Dim filesys, folder, fso, fileObj
       
    ' pega o nome do pacote, que deve ser o nome da pasta
    Set filesys = CreateObject("Scripting.FileSystemObject")
    Set folder = filesys.GetFolder(strPath) 
    pkgName = folder.Name
    
    ' inicia variaves
    ctrlFilePath = "\patches\SdmPkgHistory.txt"
    control = 0
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' verifica se o arquivo de controle existe
    ' se existe, verifica se o pacote jah foi aplicado
    If fso.FileExists(strNxroot & ctrlFilePath) Then
        Set fileObj = fso.GetFile(strNxroot & ctrlFilePath)
        Set FileContents = fileObj.OpenAsTextStream(1,-2)
        ' le linha por linha
        Do While FileContents.AtEndOfStream <> True
            LineParts = FileContents.readline
            LinePart = Split(LineParts,"@")
            ' se achar o nome do pacote, marca a variavel de controle
            If LinePart(0) = pkgName Then
                control = 1
        	End If
        loop
        ' sai do script, pois este pacote jah foi aplicado
        If control = 1 Then
            wscript.quit
        Else
        ' pacote ainda nao foi aplicado, entao registra o nome dele no arquivo de controle
            Set fileObj = fso.OpenTextFile(strNxroot & ctrlFilePath, 8, true)
            fileObj.writeLine pkgName & "@" & now()
        End If
        FileContents.close
    Else
        ' sem controle no sistema, cria o arquivo e registra o nome deste pacote
        Set oPkgFile = fso.CreateTextFile(strNxroot & ctrlFilePath, True)
        oPkgFile.writeLine "Arquivo controle dos pacotes aplicados neste sistema - " & now()
        oPkgFile.writeLine pkgName & "@" & now()
    End If


end function