Option Explicit
'***********************************************************************************************************************
'Define as variaveis de execucao do script de atualizacao do Service Desk
Dim fso, ows, ocmds, oNxEnv
Dim strNxroot, strSite, strPath, strLogPath, msg, iTimeOut, strProvider, sLinha
Set fso = CreateObject("Scripting.FileSystemObject")
Set ows = CreateObject("WScript.Shell")

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

strSite = strNxroot & "\site"
strPath = replace(WScript.ScriptFullName, WScript.ScriptName, "")
strPath = mid(strPath, 1, len(strPath) - 7)
strLogPath = strPath   
iTimeOut = 120
Set ocmds = fso.OpenTextFile(strPath & "\uteis\temp\cmds.log", 8, True)

copyFiles

msgbox "Esquema de banco de dados e objeto salvo com sucesso.", 1, "Operacao efetuada com sucesso"        


set ocmds = Nothing
Set ows = nothing
Set fso = nothing

'***********************************************************************************************************************
' Funcao: copyFiles()
' Autor: Henrique Grammelsbacher
' Data: 8/31/2010 11:08:20 PM
' Description: Copia os arquivos de nxroot para a pasta do Service Desk
Function copyFiles()
    Dim cmd
    
    
    cmd = "copy " & """" & strNxroot & "\site\mods\wsp.altertbl" & """" & " " & """" & strPath & "\nxroot\site\mods\wsp.altertbl"" /y"
    runCmdwnLog cmd

    cmd = "copy " & """" & strNxroot & "\site\mods\wsp.altercol" & """" & " " & """" & strPath & "\nxroot\site\mods\wsp.altercol"" /y"
    runCmdwnLog cmd

    cmd = "copy " & """" & strNxroot & "\site\mods\wsp_index.sch" & """" & " " & """" & strPath & "\nxroot\site\mods\wsp_index.sch"" /y"
    runCmdwnLog cmd

    cmd = "copy " & """" & strNxroot & "\site\mods\wsp_schema.sch" & """" & " " & """" & strPath & "\nxroot\site\mods\wsp_schema.sch"" /y"
    runCmdwnLog cmd

    cmd = "copy " & """" & strNxroot & "\site\mods\majic\wsp.mods" & """" & " " & """" & strPath & "\nxroot\site\mods\majic\wsp.mods"" /y"
    runCmdwnLog cmd

End Function

'***********************************************************************************************************************
' Funcao: runCmd()
' Autor: Henrique Grammelsbacher
' Data: 6/5/2009 3:03:22 PM
' Description: Executa comando e cria log do output
Function runCmd(cmd)
    runCmdwLog cmd, cmd
End Function

Function runCmdwnLog(cmd)
    Dim msg
    msg = ows.run ("cmd /c " & cmd, 0, true) 
    ocmds.writeline "EXEC - " & now & " - " & msg & " - " & cmd
    WScript.StdOut.WriteLine(">> " & cmd)    

    If ( msg > 0 and instr(cmd, "robocopy") <= 0 )  OR ( msg > 4 and instr(cmd, "robocopy") > 0 ) Then
        If Not continua(cmd) Then
            wscript.quit
        End If 
    End If 
End Function

Function runCmdwLog(cmd, logfile)
    Dim re, hora, errLog, exeLog, owss

    Set re = New RegExp
    re.pattern = " "
    re.global = true
    
    hora = retornaData
    errLog = strPath & "\uteis\temp\" &  re.replace(logfile, "_") & ".err"
    exeLog = strPath & "\uteis\temp\" &  re.replace(logfile, "_") & ".log"
    
    'msgbox "cmd /c " & cmd &  " > " & exeLog & " 2> " & errLog
    msg = ows.run ("cmd /c " & cmd &  " > " & exeLog & " 2> " & errLog ,0, true)
    ocmds.writeline "EXEC - " & now & " - " & msg & " - " & cmd
    WScript.StdOut.WriteLine(">> " & cmd)    
    
    Set re = nothing

    If ( msg > 0 and instr(cmd, "robocopy") <= 0 )  OR ( msg > 4 and instr(cmd, "robocopy") > 0 ) Then
        If Not continua(cmd) Then
            wscript.quit
        End If 
    End If 
    
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
 
