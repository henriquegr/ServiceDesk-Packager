@echo off

rem testes basicos para ver se temos um nome de pasta, estamos no dir correto e se a versao jah existe
if NOT exist "%~dp0\grep.exe" goto erro_path
if "%1"=="" goto erro_param
if exist "%~dp0\..\Userload\%1" goto erro_existe

rem criando a estrutura
mkdir ..\Userload\%1 
mkdir ..\sed\%1
mkdir ..\deref\%1

copy ..\Userload\info.txt ..\Userload\%1\info.txt 
copy ..\deref\info.txt ..\deref\%1\info.txt
copy ..\sed\info.txt ..\sed\%1\info.txt
goto end


rem mensagens de erro
:erro_existe
echo ***************************************************
echo A versao %1 jah existe
echo ***************************************************
goto end

:erro_path
echo ***************************************************
echo Este comando deve ser executado na pasta uteis
echo ***************************************************
goto end

:erro_param
echo ***************************************************
echo Deve-se informar a versao do pacote
echo ***************************************************
goto end


:end