option explicit
on error goto 0

' следующие две строки для вывода отладочных сообщений
' эти две строки можно и удалить
Dim DebugFlag 'обязательно глобальная переменная
' DebugFlag = True 'Разрешаю вывод отладочных сообщений

  Dim wshShell
  Dim fso 'as FileSystemObject
  
  Dim LogFile 'as File
  Dim sLogFile 'as string
  Dim ResDict 'as Dictionary 
  Dim ConfDict 'as Dictionary 

  Dim strCOMConnector ' as string
  Dim sFullClusterName ' as string

  Dim ServerName
  Dim KlasterPortNumber
  Dim InfoBaseName

  Dim sFullServerName

  Dim ServerName83
    Dim KlasterPortNumber83
    Dim InfoBaseName83

	Dim sFullServerName83

    Dim RepositoryPath

    Dim ClasterAdminName
    Dim ClasterAdminPass
    Dim InfoBasesAdminName
    Dim InfoBasesAdminPass
    Dim RepositoryAdminName
    Dim RepositoryAdminPass

    Dim NetFile

    Dim Folder
    Dim CountDB
    Dim Prefix
    Dim Out

	Dim NeedUpdateFromStorage
    Dim NeedDumpIB
    Dim NeedCopyFiles
    Dim NeedTestIB
    Dim NeedRestartAgent
    Dim NeedRestoreIB
    Dim NeedRestoreIB83
    Dim NeedStartIB
        
    Dim IBFile
    Dim LockMessageText
    Dim LockPermissionCode
    Dim AuthStr
    Dim TimeSleep

    Dim TimeSleepShort
    Dim Cfg
    Dim InfoCfgFile
    Dim v8exe
	Dim v83exe

    Dim TimeBeginLock
    Dim TimeEndLock

Wscript.Quit( main() )

'********************************************************************
' Возвращает 0 при успехе, 1 - при неудаче
Function main( )
    main = 1
    
  'Make sure the host is csript, if not then abort
  VerifyHostIsCscript()
  
  ' проверить версию Windows Script Host
  if CDbl(replace(WScript.Version,".",","))<5.6 then
    Echo "Для работы сценария требуется Windows Script Host версии 5.6 и выше !"
    Exit Function
  end if  

	' Инициализация сценария
	if not Init() then
	  Exit Function
	end if

	LoadSettings()

    Dim bSuccess
    bSuccess = OpenLogFile
    if not bSuccess then
		Echo(CStr(Now) + " Не удалось изменить лог-файл. лог-файл заблокирован другой программой")
		Exit Function
    end if

    Debug "ServerName", ServerName

	Echo(CStr(Now) + " НАЧАЛО ОБНОВЛЕНИЯ КОНФИГУРАЦИИ")

	Dim FindInfoBase
    FindInfoBase = DisableConnections(ServerName, ClasterAdminName, ClasterAdminPass, InfoBasesAdminName, InfoBasesAdminPass, InfoBaseName, TimeSleep)

    If NeedRestartAgent Then
        RestartAgent TimeSleep, TimeSleepShort
    End If

    Dim strCommLine, sTempFile, LineExe

    If FindInfoBase Then

        'Покажем свободное место на диске с исполняемым файлом 1С
        Echo(CStr(Now) + " " + ShowFreeSpace(v8exe))
        'Покажем свободное место на диске с архивами
        Echo(CStr(Now) + " " + ShowFreeSpace(Folder))
        
		If NeedRestoreIB Then
			Echo(CStr(Now) + " Восстановление эталонной базы")

			strCommLine = " /RestoreIB """ + IBFile + """"

			sTempFile = FSO.GetSpecialFolder(2) + "\" +FSO.GetTempName()

			LineExe = """" + v8exe + """ DESIGNER /S""" + sFullServerName + "\" + InfoBaseName + """ /UC""" + LockPermissionCode + """ /DisableStartupMessages " + AuthStr + " " + strCommLine + " /Out""" + sTempFile +""""
			Echo(CStr(Now) + " ком.строка запуска: " + LineExe)

			wshShell.Run LineExe, 5, True

			Show1CConfigLog sTempFile, " Ошибка при загрузке базы из файла"
		End If

		if NeedUpdateFromStorage then
			Echo(CStr(Now) + " Обновление конфигурации из хранилища")
			
			Dim strRepository, UpdateFromStorage
			strRepository = " /ConfigurationRepositoryF"""+RepositoryPath+""""
			strRepository = strRepository + " /ConfigurationRepositoryN"""+RepositoryAdminName + """ /ConfigurationRepositoryP"""+RepositoryAdminPass+""""
			
			UpdateFromStorage = " /ConfigurationRepositoryUpdateCfg -v -force -revised " ' обновляем из хранилища
			' /LoadCfg — загрузка конфигурации из файла; 
			' /UpdateCfg — обновление конфигурации, находящейся на поддержке; 
			' /ConfigurationRepositoryUpdateCfg — обновление конфигурации из хранилища; 
			' /LoadConfigFiles — загрузить файлы конфигурации.

			sTempFile = FSO.GetSpecialFolder(2) + "\" +FSO.GetTempName()
			
			LineExe = """" + v8exe + """ DESIGNER /S""" + sFullServerName + "\" + InfoBaseName + """ /UC""" + LockPermissionCode + """ /DisableStartupMessages " + AuthStr + " " +UpdateFromStorage + strRepository + " /Out""" + sTempFile +""""
			Echo(CStr(Now) + " ком.строка запуска: "+LineExe)
			'LogFile.Close()
			'LogFile = ""

			' Обновим конфигурацию из хранилища
			wshShell.Run LineExe, 5, True

			' анализирую лог работы конфигуратора, т.к могут быть ошибки, например, Ошибка при выполнении операции с информационной базой или Ошибка обновления конфигурации из хранилища
			'или Для выполнения операции требуется получение объектов:
			'или  Операция с хранилищем конфигурации отменена
			' также лог конфигуратора показываю в своем логе
			Show1CConfigLog sTempFile, " Ошибка при обновлении конфигурации из хранилища"
		end if ' NeedUpdateFromStorage
		
        Echo(CStr(Now) + " Обновление конфигурации ИБ") 'EchoWithOpenAndCloseLog

		sTempFile = FSO.GetSpecialFolder(2) + "\" +FSO.GetTempName()

        LineExe = """" + v8exe + """ DESIGNER /S""" + sFullServerName + "\" + InfoBaseName + """ /UC""" + LockPermissionCode + """ /DisableStartupMessages " + AuthStr + " " + " /UpdateDBCfg -Server /Out""" + sTempFile + """ -NoTruncate"
        Echo(CStr(Now) + " ком.строка запуска: "+LineExe) ' EchoWithOpenAndCloseLog

        ' Обновим конфигурацию БД
        wshShell.Run LineExe, 5, True

		' анализирую лог работы конфигуратора, т.к могут быть ошибки и база не обновится
		' также лог конфигуратора показываю в своем логе
		Show1CConfigLog sTempFile, " Ошибка при обновлении базы данных"
		
        If NeedTestIB = True Then
            Echo(CStr(Now) + " тестируем базу и пересчитываем итоги.") ' EchoWithOpenAndCloseLog

			sTempFile = FSO.GetSpecialFolder(2) + "\" +FSO.GetTempName()
			
            LineExe = """" + v8exe + """ DESIGNER /S""" + sFullServerName + "\" + InfoBaseName + """ /UC""" + LockPermissionCode + """ /DisableStartupMessages " + AuthStr + " /IBCheckAndRepair -LogIntegrity -RecalcTotals /Out""" + sTempFile + """ -NoTruncate"
            Echo(CStr(Now) + " ком.строка: " + LineExe) ' EchoWithOpenAndCloseLog

            wshShell.Run LineExe, 5, True

			Show1CConfigLog sTempFile, " Ошибка при выгрузке базы данных"
        End if
        
        if NeedStartIB then
            Echo(CStr(Now) + " обновляем базу в режиме Предприятия.") ' EchoWithOpenAndCloseLog

			sTempFile = FSO.GetSpecialFolder(2) + "\" +FSO.GetTempName()

            'EnableConnections ServerName, ClasterAdminName, ClasterAdminPass, InfoBasesAdminName, InfoBasesAdminPass, InfoBaseName, True

            LineExe = """" + v8exe + """ ENTERPRISE /CCLOSE /S""" + sFullServerName + "\" + InfoBaseName + """ /UC""" + LockPermissionCode + """ /DisableStartupMessages " + AuthStr + " /Out""" + sTempFile + """ -NoTruncate"
            Echo(CStr(Now) + " ком.строка: " + LineExe) ' EchoWithOpenAndCloseLog

            wshShell.Run LineExe, 5, True

    		'DisableConnections ServerName, ClasterAdminName, ClasterAdminPass, InfoBasesAdminName, InfoBasesAdminPass, InfoBaseName, TimeSleepShort

			Show1CConfigLog sTempFile, " Ошибка при обновлении базы в режиме Предприятия"
        end if

        If FSO.FolderExists(Folder) = False Then
            FSO.CreateFolder Folder
        End if
        
        If NeedDumpIB = True Then
            Echo(CStr(Now) + " выгружаем базу данных в архив") ' EchoWithOpenAndCloseLog
			
			Dim formatDate
			formatDate = GetFormatDay
			sTempFile = FSO.GetSpecialFolder(2) + "\" +FSO.GetTempName()

            LineExe = """" + v8exe + """ DESIGNER /S""" + sFullServerName + "\" + InfoBaseName + """ /UC""" + LockPermissionCode + """ /DisableStartupMessages " + AuthStr + " /DumpIB""" + Folder + Prefix + formatDate + ".dt"" /Out""" + sTempFile + """ -NoTruncate"
            Echo(CStr(Now) + " ком.строка: " + LineExe) ' EchoWithOpenAndCloseLog

            wshShell.Run LineExe, 5, True

            Dim haveProblem
			haveProblem = Show1CConfigLog(sTempFile, " Ошибка при выгрузке базы данных")
			If Not haveProblem And NeedRestoreIB83 Then
    			Echo(CStr(Now) + " Восстановление базы в 8.3")

    			strCommLine = " /RestoreIB """ + Folder + Prefix + formatDate + ".dt"""

    			sTempFile = FSO.GetSpecialFolder(2) + "\" +FSO.GetTempName()

    			LineExe = """" + v83exe + """ DESIGNER /S """ + sFullServerName83 + "\" + InfoBaseName83 + """ /UC """ + LockPermissionCode + """ /DisableStartupMessages " + AuthStr + " " + strCommLine + " /Out """ + sTempFile +""""
    			Echo(CStr(Now) + " ком.строка запуска: " + LineExe)

    			wshShell.Run LineExe, 5, True

    			Show1CConfigLog sTempFile, " Ошибка при загрузке базы из файла"
			End If
        End if
		
        'OpenLogFile

        Echo(CStr(Now) + " Установка разрешения подключения к ИБ")

        FindInfoBase = False
        
        EnableConnections ServerName, ClasterAdminName, ClasterAdminPass, InfoBasesAdminName, InfoBasesAdminPass, InfoBaseName, False

    End If ' FindInfoBase

    WriteLogIntoIBEventLog sFullServerName, InfoBaseName, sLogFile

    If NeedCopyFiles = True Then 
        If fso.FileExists(NetFile) Then
            fso.DeleteFile(NetFile)
        End If
        Debug "Out", Out
		Debug "NetFile", NetFile

        fso.MoveFile Out, NetFile
        
        Debug "NetFile", NetFile
    End if

    If NeedDumpIB = True Then 
        CALL DelOldFiles(Folder, CountDB)
    End if

    main = 0
End Function

Function Show1CConfigLog(sTempFile, errorMessage)
	Dim configLogFile, haveProblem, errorString, errorPos

	Set configLogFile = fso.OpenTextFile(sTempFile, 1)

	haveProblem = false
	Do While configLogFile.AtEndOfStream <> True
		errorString = configLogFile.ReadLine
		Echo errorString
		errorPos = InStr(1, errorString, "Ошибка", 1)
		If errorPos > 0 Then
			haveProblem = true
		end if
	Loop
	if haveProblem = true then
		Echo(CStr(Now) + errorMessage) ' EchoWithOpenAndCloseLog '" Ошибка при обновлении конфигурации из хранилища")
	end if
	configLogFile.Close()

	Show1CConfigLog = haveProblem
End Function

Sub WriteLogIntoIBEventLog(sFullServerName, InfoBaseName, sLogFile)
		'Sub WriteLogIntoIBEventLog(ServerName, KlasterPortNumber, InfoBaseName, sLogFile)
    Echo(CStr(Now) + " Сохранение лога в журнал регистрации ИБ")
    Dim ComConnector, connection

    Set ComConnector = CreateCOMConnector() ' CreateObject("v82.COMConnector")
        'Set connection = ComConnector.Connect("Srvr=" + ServerName + ":" + CStr(KlasterPortNumber) + ";Ref=" + InfoBaseName + ";Usr=" + InfoBasesAdminName + ";Pwd=" + InfoBasesAdminPass)
    Set connection = ComConnector.Connect("Srvr=" + sFullServerName + ";Ref=" + InfoBaseName)

    Echo(CStr(Now) + " ЗАВЕРШЕНИЕ ОБНОВЛЕНИЯ КОНФИГУРАЦИИ")

    'LogFile.Close()
    'LogFile = ""

    Dim f, Text
    Set f = fso.OpenTextFile(sLogFile, 1, False, -2) 'Out
    Text = f.ReadAll

    'Запишем всю информацию из log-файла в журнал регистрации
    connection.WriteLogEvent "Регламентное обновление ИБ", connection.EventLogLevel.Information,,, Text

    connection = Null
    ComConnector = Null
    f = Null
End Sub

Function CreateCOMConnector()
    Echo(CStr(Now) + " Создание COM-коннектора <"+ strCOMConnector + ">")
    Dim ComConnector
    Set ComConnector = CreateObject(strCOMConnector) ' CreateObject("v82.COMConnector")

    set CreateCOMConnector = ComConnector
End Function

Function DisableConnections(ServerName, ClasterAdminName, ClasterAdminPass, InfoBasesAdminName, InfoBasesAdminPass, InfoBaseName, TimeSleep)
    'Echo(CStr(Now) + " Создание COM-коннектора")
    Dim ComConnector
    Set ComConnector = CreateCOMConnector() ' CreateObject("v82.COMConnector")

    Echo(CStr(Now) + " Подключение к агенту сервера")
    Dim ServerAgent
    Set ServerAgent = ComConnector.ConnectAgent(sFullClusterName) ' ComConnector.ConnectAgent(ServerName)

    Echo(CStr(Now) + " Получение массива кластеров сервера у агента сервера")
    Dim Clasters
    Clasters = ServerAgent.GetClusters()

    Echo(CStr(Now) + " Начало завершения работы пользователей")

    Echo(CStr(Now) + " Начало цикла нахождения необходимого кластера по имени")
    Dim findClaster, i, Claster
    findClaster = false
    For i = LBound(Clasters) To UBound(Clasters)
                'If Claster.MainPort = KlasterPortNumber Then
        Set Claster = Clasters(i)

        if (UCase(Claster.HostName) = UCase(ServerName)) then
            findClaster = true
            Exit for
        End if
    Next
    if findClaster = false then
        Echo(CStr(Now) + " Ошибка - не нашли кластер <"+sFullClusterName+">") 'ServerName
        Exit Function
    end if
            
    Echo(CStr(Now) + " Аутентикация к найденному кластеру: " + Claster.ClusterName + ", "+Claster.HostName)
	    'Echo(CStr(Now) + " Аутентикация к найденному кластеру: " + Claster.Name + ", "+Claster.HostName)
 
    ServerAgent.Authenticate Claster, ClasterAdminName, ClasterAdminPass

    Echo(CStr(Now) + " Получение списка работающих рабочих процессов и обход в цикле")

    Dim WorkServers, WorkServer, WorkingProcesses
    WorkServers = ServerAgent.GetWorkingServers(Claster)
    For i = LBound(WorkServers) To UBound(WorkServers)
        Set WorkServer = WorkServers(i)
        Echo(CStr(Now) + " Обрабатываю рабочий сервер "+WorkServer.Name+": " + WorkServer.HostName)

        WorkingProcesses = ServerAgent.GetServerWorkingProcesses(Claster, WorkServer)
			'set WorkingProcesses = ServerAgent.GetWorkingProcesses(Claster, WorkServer)
			'Если РабочиеПроцессы <> НЕОПРЕДЕЛЕНО Тогда

		Dim FindInfoBase, j, ConnectToWorkProcess
		FindInfoBase = false

        For j = LBound(WorkingProcesses) To UBound(WorkingProcesses)

            If WorkingProcesses(j).Running = 1 Then

                Echo(CStr(Now) + " Создание соединения с рабочим процессом " + WorkingProcesses(j).HostName + ":" + CStr(WorkingProcesses(j).MainPort))
                Set ConnectToWorkProcess = ComConnector.ConnectWorkingProcess("tcp://" + WorkingProcesses(j).HostName + ":" + CStr(WorkingProcesses(j).MainPort))

                ConnectToWorkProcess.AuthenticateAdmin ClasterAdminName, ClasterAdminPass
                ConnectToWorkProcess.AddAuthentication InfoBasesAdminName, InfoBasesAdminPass

                If Not FindInfoBase Then

                    Echo(CStr(Now) + " Получение списка ИБ рабочего процесса")
                    Dim InfoBases, InfoBase
                    InfoBases = ConnectToWorkProcess.GetInfoBases()

                    Echo(CStr(Now) + " Поиск нужной ИБ")
                    For h = LBound(InfoBases) To UBound(InfoBases)
                        Echo(CStr(Now) + " Обрабатывается ИБ: " + InfoBases(h).Name)
                        If LCase(InfoBases(h).Name) = LCase(InfoBaseName) Then
                            Set InfoBase = InfoBases(h)
                            FindInfoBase = True
                            Echo(CStr(Now) + " Нашли нужную ИБ")
                            Exit For
                        End If
                    Next

                    If Not FindInfoBase Then
                        Echo(CStr(Now) + " Не нашли нужную ИБ <"+InfoBaseName+">")
                        Exit Function
                    End If

                    Echo(CStr(Now) + " Установка запрета на подключения к ИБ: " + InfoBase.Name)
                    InfoBase.ConnectDenied = True
                    InfoBase.ScheduledJobsDenied = True
                    InfoBase.DeniedFrom = TimeBeginLock
                    InfoBase.DeniedTo   = TimeEndLock
                    InfoBase.DeniedMessage = LockMessageText
                    InfoBase.PermissionCode = LockPermissionCode
                    ConnectToWorkProcess.UpdateInfoBase(InfoBase)

                    InfoBases = ServerAgent.GetInfoBases(Claster)

                    Echo(CStr(Now) + " Поиск нужной ИБ для сессии")
                    Dim h, InfoBaseSession
                    For h = LBound(InfoBases) To UBound(InfoBases)
                        Echo(CStr(Now) + " Обрабатывается ИБ: " + InfoBases(h).Name)
                        If LCase(InfoBases(h).Name) = LCase(InfoBaseName) Then
                            Set InfoBaseSession = InfoBases(h)
								'FindInfoBase = True
                            Echo(CStr(Now) + " Нашли нужную ИБ для сессии")
                            Exit For
                        End If
                    Next

                    ' Устанавливаем задержку выполнения
                    Echo(CStr(Now) + " Задержка перед началом завершения работы пользователей "+CStr(TimeSleep/1000) + " секунд")
                    'Echo(CStr(Now) + " Задержка перед началом завершения работы пользователей")
                    'set WshShell = WScript.CreateObject("WScript.Shell")
                    WScript.Sleep TimeSleep 

                End If

                If FindInfoBase Then
	                Echo(CStr(Now) + " Начало завершение работы пользователей с ИБ " + InfoBase.Name)

                    Echo(CStr(Now) + " Обработка списка сеансов")
                    Dim Sessions, Session, k, UserName, AppID
                    Sessions = ServerAgent.GetInfoBaseSessions(Claster, InfoBaseSession)
                    For k = LBound(Sessions) To UBound(Sessions)
                        Set Session = Sessions(k)
                        UserName = Session.UserName
							'ConnID    = Session.ConnID;
                        AppID    = UCase(Session.AppID)        
                        
							'Если НЕ отключаемКонфигуратор И нРег(AppID) = "designer" Тогда
							'//Если нРег(AppID) = "backgroundjob" ИЛИ нРег(AppID) = "designer" Тогда
							'    // если это сеансы конфигуратора или фонового задания, то не отключаем
							'    Продолжить;
							'КонецЕсли;
							'//Если UserName = ИмяПользователя() Тогда
							'//    // это текущий пользователь
							'//    Продолжить;
							'//КонецЕсли;
                        Echo(CStr(Now) + " Отключено соединение: " + "User=["+UserName+"] ConnID=["+""+"] AppID=["+AppID+"]")
                        ServerAgent.TerminateSession Claster, Session
                    next

                    'if false then
                    '    Echo(CStr(Now) + " Обработка списка соединений")
                    '    Connections = ConnectToWorkProcess.GetInfoBaseConnections(InfoBase)
                    '    For k = LBound(Connections) To UBound(Connections)
                    '        Echo(CStr(Now) + " Разрываем соединение: Пользователь " + Connections(k).UserName + ", компьютер " + Connections(k).HostName + ", установлено " + CStr(Connections(k).ConnectedAt) + ", режим " + Connections(k).AppID)
                    '        If Connections(k).AppID = "SrvrConsole" Then
                    '            ' Не трогаем соединения консоли, оно никому не мешает
                    '        ElseIf Connections(k).AppID = "COMConsole" Then
                    '            ' Не трогаем соединения консоли, оно никому не мешает
                    '        Else
                    '            ConnectToWorkProcess.Disconnect(Connections(k))
                    '            Echo(CStr(Now) + " Отключено соединение: Пользователь " + Connections(k).UserName + ", компьютер " + Connections(k).HostName + ", установлено " + CStr(Connections(k).ConnectedAt) + ", режим " + Connections(k).AppID)
                    '        End If
                    '    Next
                    'End If

	                Echo(CStr(Now) + " Окончание завершения работы пользователей")
                End If

            End If
        Next
    Next

    ComConnector = Null
    ServerAgent = Null
    Clasters = Null
    WorkingProcesses = Null
    ConnectToWorkProcess = Null
    InfoBases = Null
    InfoBase = Null
    'Connections = Null

	DisableConnections = FindInfoBase
End Function

Function EnableConnections(ServerName, ClasterAdminName, ClasterAdminPass, InfoBasesAdminName, InfoBasesAdminPass, InfoBaseName, ConnectionsOnly)
    EnableConnections = false
    
    Dim ComConnector, ServerAgent, Clasters, Claster, findClaster, i, j, h

    Set ComConnector = CreateCOMConnector() ' CreateObject("v82.COMConnector")
    Set ServerAgent = ComConnector.ConnectAgent(sFullClusterName) 'ServerName)
    Clasters = ServerAgent.GetClusters()

    findClaster = false
    For i = LBound(Clasters) To UBound(Clasters)
    
                'If Claster.MainPort = KlasterPortNumber Then
        Set Claster = Clasters(i)
        if (UCase(Claster.HostName) = UCase(ServerName)) then
            findClaster = true
            Exit for
        End if
    Next
    if findClaster = false then
        Echo(CStr(Now) + " Ошибка - не нашли кластер "+sFullClusterName) 'ServerName
        Exit Function
    end if

    ServerAgent.Authenticate Claster, ClasterAdminName, ClasterAdminPass

    'WorkingProcesses = ServerAgent.GetWorkingProcesses(Claster)
    
    Dim WorkServers, WorkServer, WorkingProcesses, ConnectToWorkProcess, FindInfoBase

    FindInfoBase = False

    WorkServers = ServerAgent.GetWorkingServers(Claster)
    For i = LBound(WorkServers) To UBound(WorkServers)
        Set WorkServer = WorkServers(i)
        Echo(CStr(Now) + " Обрабатываю рабочий сервер "+WorkServer.Name+": " + WorkServer.HostName)

        WorkingProcesses = ServerAgent.GetServerWorkingProcesses(Claster, WorkServer)

        For j = LBound(WorkingProcesses) To UBound(WorkingProcesses)

            If WorkingProcesses(j).Running = 1 Then

                Set ConnectToWorkProcess = ComConnector.ConnectWorkingProcess("tcp://" + WorkingProcesses(j).HostName + ":" + CStr(WorkingProcesses(j).MainPort))
                ConnectToWorkProcess.AuthenticateAdmin ClasterAdminName, ClasterAdminPass
                ConnectToWorkProcess.AddAuthentication InfoBasesAdminName, InfoBasesAdminPass

                ' Получаем список ИБ рабочего процесса
                Dim InfoBases, InfoBase
                InfoBases = ConnectToWorkProcess.GetInfoBases()
                For h = LBound(InfoBases) To UBound(InfoBases)
                    If LCase(InfoBases(h).Name) = LCase(InfoBaseName) Then
                        Set InfoBase = InfoBases(h)
                        FindInfoBase = True
                        Exit For
                    End If
                Next

                If FindInfoBase Then
                    ' Устанавливаем разрешение на подключение соединений
                    InfoBase.ConnectDenied = False
                    If Not ConnectionsOnly Then
                        InfoBase.ScheduledJobsDenied = false
                    End If
                    InfoBase.DeniedMessage = ""
                    InfoBase.PermissionCode = ""
                    ConnectToWorkProcess.UpdateInfoBase(InfoBase)
                    Exit For
                End If
                
                if not FindInfoBase then
                    Echo(CStr(Now) + " Ошибка - не нашли базу для снятия запрета на подключение соединений <"+InfoBaseName+">")
                    Exit Function
                end if

            End If

        Next
    Next

    EnableConnections = true
End Function

Sub RestartAgent(TimeSleep, TimeSleepShort)
    Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
    
    'Stop Service
    strServiceName = "1C:Enterprise 8.2 Server Agent" ' "1C:Enterprise 8.1 Server Agent"
    Set colListOfServices = objWMIService.ExecQuery("Select * from Win32_Service Where Name ='" & strServiceName & "'")
    For Each objService in colListOfServices
        objService.StopService()
        Echo(CStr(Now) + " " + CStr(objService.Name) + " Остановка службы сервера 1С Предприятия")
    Next
        
    WScript.Sleep TimeSleep
        
    TerminateProcess "ragent.exe"
    TerminateProcess "rmngr.exe"
    TerminateProcess "rphost.exe"
                    
    WScript.Sleep TimeSleepShort
    
    'Start Service
    strServiceName = "1C:Enterprise 8.2 Server Agent" ' "1C:Enterprise 8.1 Server Agent"
        'Set colListOfServices = objWMIService.ExecQuery ("Select * from Win32_Service Where Name ='" & strServiceName & "'")
    For Each objService in colListOfServices
        objService.StartService()
        Echo(CStr(Now) + " " + CStr(objService.Name) + " Запуск службы сервера 1С Предприятия")
    Next
        
    WScript.Sleep TimeSleepShort
End Sub

Sub TerminateProcess(strProcessName)
    Set colProcess = objWMIService.ExecQuery ("Select * from Win32_Process Where Name = '" & strProcessName & "'")
    For Each objProcess in colProcess
        objProcess.Terminate()
        Echo(CStr(Now) + " " + CStr(objProcess.Name) + " Завершение процесса агента сервера 1С Предприятия")
    Next
End Sub

Sub LoadSettings()
    ServerName = ResDict.Item(LCase("ServerName")) ' "WorkServer" 'Имя сервера БД
    KlasterPortNumber = ResDict.Item(LCase("KlasterPortNumber")) ' 1541 'Номер пора кластера
    InfoBaseName = ResDict.Item(LCase("InfoBaseName")) ' "myBase_User_01" 'Имя ИБ

	sFullServerName = ServerName
	sFullClusterName = ServerName
	if "" <> CStr(KlasterPortNumber) then
		sFullServerName = ServerName + ":" + CStr(KlasterPortNumber)
		sFullClusterName = ServerName + ":" + CStr(CInt(KlasterPortNumber)-1)
	end if

    ServerName83 = ResDict.Item(LCase("ServerName83")) 
    KlasterPortNumber83 = ResDict.Item(LCase("KlasterPortNumber83")) 
    InfoBaseName83 = ResDict.Item(LCase("InfoBaseName83")) 

	sFullServerName83 = ServerName83
	if "" <> CStr(KlasterPortNumber83) then
		sFullServerName83 = sFullServerName83 + ":" + CStr(KlasterPortNumber83)
	end if

    RepositoryPath = ResDict.Item(LCase("RepositoryPath")) ' "E:\Repository\test" ' путь к хранилищу

		' ВАЖНО - база должна быть ранее зарегистрирована как server:port\baseName - например, WorkServer:1541\myBase_User_01
		' если база была зарегистрирована как server\baseName (без указания порта) - при работе с хранилищем будет ошибка из-за измененения местонахождения информационной базы
		' если такой регистрации не было, не удается загрузить изменения из хранилища

    ' ВАЖНО не должен быть запущен процесс Конфигуратора с подключением к хранилищу с таким пользователем !
    ' будет ошибка - (Пользователь уже аутентифицирован в хранилище.)

    ClasterAdminName = ResDict.Item(LCase("ClasterAdminName")) ' "" 'Имя администратора кластера
    ClasterAdminPass = ResDict.Item(LCase("ClasterAdminPass")) ' "" 'Пароль администратора кластера
    InfoBasesAdminName = ResDict.Item(LCase("InfoBasesAdminName")) ' "Администратор1" 'Имя администратора ИБ
    InfoBasesAdminPass = ResDict.Item(LCase("InfoBasesAdminPass")) ' "" 'Пароль администратора ИБ
    RepositoryAdminName = ResDict.Item(LCase("RepositoryAdminName")) ' "Администратор" ' Имя администратора хранилища
    RepositoryAdminPass = ResDict.Item(LCase("RepositoryAdminPass")) ' "" ' Пароль администратора хранилища

		'FilePath = ResDict.Item(LCase("FilePath")) ' "\\WorkServer\Share\Admin1C\confupdate.vbs" 'Путь к текущему файлу
    NetFile = ResDict.Item(LCase("NetFile")) ' "\\WorkServer\Share\Admin1C\confupdate_base.txt" 'Путь к log-файлу в сети - используется только для NeedCopyFiles = True

    Folder = ResDict.Item(LCase("Folder")) ' "\\WorkServer\Share\Admin1C\" 'Каталог для выгрузки базы
    CountDB = CInt(ResDict.Item(LCase("CountDB"))) ' 7 'За сколько дней хранить копии
    Prefix = ResDict.Item(LCase("Prefix")) ' "base" 'Префикс файла выгрузки
    Out = ResDict.Item(LCase(LCase("LogFile"))) ' "\\WorkServer\Share\Admin1C\confupdate.txt" 'Путь к log-файлу
    sLogFile = Out
    Debug "Out", Out

		'UpdateFromStorage = ResDict.Item(LCase("UpdateFromStorage")) ' " /ConfigurationRepositoryUpdateCfg -v -force -revised " ' обновляем из хранилища

	NeedUpdateFromStorage = UCase(ResDict.Item(LCase("NeedUpdateFromStorage"))) = "TRUE" ' Необходимость обновления конфигурации из хранилища конфигурации
    NeedDumpIB = UCase(ResDict.Item(LCase("NeedDumpIB"))) = "TRUE" ' True ' Необходимость выгрузки базы
    NeedCopyFiles = UCase(ResDict.Item(LCase("NeedCopyFiles"))) = "TRUE" ' True ' Необходимость выгрузки базы
    NeedTestIB = UCase(ResDict.Item(LCase("NeedTestIB"))) = "TRUE" ' False ' Необходимость тестирования базы
    NeedRestartAgent = UCase(ResDict.Item(LCase("NeedRestartAgent"))) = "TRUE" ' False ' Необходимость рестарта агента сервера
    NeedRestoreIB = UCase(ResDict.Item(LCase("NeedRestoreIB"))) = "TRUE" ' Необходимость восстановления конфигурации из файла
    NeedRestoreIB83 = UCase(ResDict.Item(LCase("NeedRestoreIB83"))) = "TRUE" ' Необходимость восстановления конфигурации из файла платформой 8.3
    NeedStartIB = UCase(ResDict.Item(LCase("NeedStartIB"))) = "TRUE" ' Необходимость запуска 1С после обновления из хранилища для обновления в режиме Предприятия
        
    IBFile = ResDict.Item(LCase("IBFile")) ' "" 'Путь к файлу с выгрузкой базы
    LockMessageText = ResDict.Item(LCase("LockMessageText")) ' "Идет регламент. Подождите..." 'Текст сообщения о блокировки подключений к ИБ
    LockPermissionCode = ResDict.Item(LCase("LockPermissionCode")) ' "Артур" 'Ключ для запуска заблокированной ИБ
    AuthStr = ResDict.Item(LCase("AuthStr")) ' "/WA+" 
    TimeSleep = CLng(ResDict.Item(LCase("TimeSleep"))) ' 10000 '600000 '10 секунд 600 секунд

    TimeSleepShort = CLng(ResDict.Item(LCase("TimeSleepShort"))) ' 2000 '60000 '2 секунд 60 секунд
    Cfg = ResDict.Item(LCase("Cfg")) ' "" 'Путь к файлу с измененной конфигурацией
    InfoCfgFile = ResDict.Item(LCase("InfoCfgFile")) ' "" 'Информация о файле обновления конфигурации
    v8exe = ResDict.Item(LCase("v8exe")) ' "C:\Program Files (x86)\1cv82\8.2.18.96\bin\1cv8.exe" 'Путь к исполняемому файлу 1С:Предприятия 8.2
	v83exe = ResDict.Item(LCase("v83exe"))
		'rem NewPass = "" 'Новый пароль администратора, обновляющего ИБ
    strCOMConnector = ResDict.Item(LCase("COMConnector"))


    TimeBeginLock = Now ' Время начала блокировки ИБ
    TimeEndLock = DateAdd("h", 2, TimeBeginLock) ' Время окончания блокировки ИБ
End Sub

' Инициализация сценария
Function Init( )
      Init = false
        
      set wshShell = wScript.createObject("wScript.shell")
      Set fso = CreateObject("Scripting.FileSystemObject") 
      
    ' задать имя ini-файла
      Dim IniFileName

      Dim intOpMode
      intOpMode = intParseCmdLine(IniFileName)

			' всегда один ini-файл в каталоге программы
			'  IniFileName = Replace(LCase(WScript.ScriptFullName),".vbs",".ini")
    Debug "IniFileName", IniFileName

      if not GetDataFromIniFile(IniFileName, ResDict) then
        Exit Function
      end if

    On Error Resume Next
      Dim sDebugFlag
      sDebugFlag = ResDict.Item(LCase("DebugFlag"))
      Debug "sDebugFlag",sDebugFlag
      if sDebugFlag<>"" then
        DebugFlag = CBool(sDebugFlag)
      end if
      Debug "DebugFlag",DebugFlag
    On Error Goto 0

        '' получить лог-файл
        '  LogFile = Null 'не выводить в лог-файл, если не задан путь к нему
        '  Dim sLogFile
        '  sLogFile = ResDict.Item(LCase("LogFile"))
        '  if sLogFile<>"" then
        '    If (NOT blnOpenFile(sLogFile, LogFile)) Then
        '      Call Wscript.Echo ("Не могу открыть лог-файл <"+sLogFile+"> .")
        '      Exit Function
        '    End If
        '  End If    

  Init = true
End Function 'Init      

Function GetFormatDay()
    iDay = Day(Now)
    mDay = CStr(Day(Now))
    iMonth = Month(Now)
    mMonth = CStr(Month(Now))
    mYear = CStr(Year(Now))

    nCDay = "_" + mYear + "_"
    If iMonth < 10 Then
       nCDay = nCDay + "0"
    End If
    nCDay = nCDay + mMonth + "_"
    If iDay < 10 Then
       nCDay = nCDay + "0"
    End If
    nCDay = nCDay + mDay
	
	GetFormatDay = nCDay
End Function

' Функция для определения свободного места на диске
Function ShowFreeSpace(drvPath)
  Dim d, s
  on error Resume next
  Set d = fso.GetDrive(fso.GetDriveName(drvPath))
  s = "Drive " & UCase(drvPath) & " - " 
  s = s & d.VolumeName  & " "
  s = s & "Free Space: " & FormatNumber(d.FreeSpace/1024/1024, 0) 
  s = s & " Mbytes"
  on error goto 0
  ShowFreeSpace = s
End Function

' Скрипт для затирания устаревших файлов: 
' Удаляет только файлы у которых сходятся префиксы
Sub DelOldFiles(Folder_Name, Stack_Depth)
    Set folder = fso.GetFolder(Folder_Name)
    Set files = folder.Files
    For Each f in files
        fdate = f.DateCreated
        fPrefix = Left(f.Name,Len(Prefix))
        If ((Date - fdate) > Stack_Depth) And fPrefix = Prefix Then
            f.Delete
        End If
    Next
End Sub

' получить данные из INI-файла
' ResDict - объект Dictionary, где хранятся пары ключ/значение
Function GetDataFromIniFile(ByVal IniFileName, ByRef ResDict)
      GetDataFromIniFile = false
  
    ' далее автоматически
    Dim IniFile 'As TextStream

    On Error Resume Next
    Dim ForRead
    ForRead =1
    Set IniFile = fso.OpenTextFile(IniFileName,ForRead)
    if err.Number<>0 then
      Err.Clear()
      echo "Ini-файл "& IniFileName &" не удалось открыть!"
      Exit Function
    end if
    on error goto 0

    Set ResDict = CreateObject("Scripting.Dictionary")
    Dim s, Matches, Match
    Dim reg 'As RegExp
    Set reg = new RegExp
      reg.Pattern="^\s*([^=]+)\s*=\s*([^;']+)[;']?"
      reg.IgnoreCase = True

    Dim elem, index, lkey, lvalue

    Do While IniFile.AtEndOfStream <> True
      s = IniFile.ReadLine
    ' если не строка-комментарий  
      if not RegExpTest("^\s*[;']",s) then
    '  For index=0 To IniDict.Count-1
    '    reg.Pattern="\s*"+elem(index)+"\s*=\s*(.+)"
    ' выделить ключ и значение в Ini-файле, убрав возможные комментарии
        Set Matches = reg.Execute(s)
        if Matches.Count>0 then
   
		' добавить новую пару, исключив из значения табуляцию и левые(и правые) пробелы    
					'ResDict.Add elem(index),Trim(replace(Matches(0).SubMatches(0),vbTab," "))
            lkey = LCase(Trim(replace(Matches(0).SubMatches(0),vbTab," ")))
            lvalue = replace(Matches(0).SubMatches(1), vbTab, " ")
            lvalue = Trim(replace(lvalue, chr(34), "")) 'убираю кавычки
            
            ResDict.Add lkey, lvalue
					'ResDict.Add LCase(Trim(replace(Matches(0).SubMatches(0),vbTab," "))),Trim(replace(Matches(0).SubMatches(1),vbTab," "))

Debug "lkey=lvalue", lkey + " = [" + lvalue + "]"
        end if
      end if
    Loop
    IniFile.Close()

    if ResDict.Count=0 then
      echo "Не удалось прочесть данные из Ini-файла " & IniFileName
      GetDataFromIniFile = false
    else  
      GetDataFromIniFile = true
    end if
End Function 'GetDataFromIniFile


' проверить на соответствие шаблону
' регистр символов не важен
  Dim regExTest               ' Create variable.
Function RegExpTest(ByVal patrn, ByVal strng)
  if IsEmpty(regExTest) then
    Set regExTest = New RegExp         ' Create regular expression.
  end if
  regExTest.Pattern = patrn         ' Set pattern.
  regExTest.IgnoreCase = true      ' disable case sensitivity.
  RegExpTest = regExTest.Test(strng)      ' Execute the search test.
'  regEx = null
End Function

Function OpenLogFile()
	Echo sLogFile

	on error resume next
    Set LogFile = fso.OpenTextFile(sLogFile, 8, True)
  
    if err.Number<>0 then
    	err.Clear()
        LogFile = nothing
	    OpenLogFile = false
	    on error goto 0
	    Exit Function
    end if
	on error goto 0

    OpenLogFile = true 'set OpenLogFile = LogFile
End Function

Sub Echo(text)
  WScript.Echo(text)
on error resume next
  If IsObject(LogFile) then        'LogFile should be a file object
    LogFile.WriteLine text
  end if
on error goto 0
End Sub'Echo

Sub EchoWithOpenAndCloseLog(text)
    OpenLogFile

    Echo(text)    

    LogFile.Close()    
    LogFile = ""
End Sub'Echo

Sub Debug(ByVal title, ByVal msg)
'exit sub
on error resume next
  DebugFlag = DebugFlag
  if err.Number<>0 then
    err.Clear()
    on error goto 0
    Exit Sub
  end if
  if DebugFlag then
    if not (IsEmpty(msg) or IsNull(msg)) then
      msg = CStr(msg)
    end if
    if not (IsEmpty(title) or IsNull(title)) then
      title = CStr(title)
    end if
    If msg="" Then
      Echo(title)
    else
      Echo(title+" - <"+msg+">")
    End If
  End If
on error goto 0
End Sub'Debug

Private Function intParseCmdLine( ByRef strFileName)

	Dim strFlag 'intParseCmdLine

'    ON ERROR RESUME NEXT
    If Wscript.Arguments.Count > 0 Then
        strFlag = Wscript.arguments.Item(0)
    End If

    If IsEmpty(strFlag) Then                'No arguments have been received
        ShowUsage 'intParseCmdLine = CONST_SHOW_USAGE
        Exit Function
    End If

        'Check if the user is asking for help or is just confused
    If (strFlag="help") OR (strFlag="/h") OR (strFlag="\h") OR (strFlag="-h") _
        OR (strFlag = "\?") OR (strFlag = "/?") OR (strFlag = "?") _
        OR (strFlag="h") Then
        ShowUsage 'intParseCmdLine = CONST_SHOW_USAGE
        Exit Function
    End If
    intParseCmdLine = 0 'CONST_LIST

    strFilename = strFlag
End Function

Sub ShowUsage ()

    Wscript.Echo ""
'    Wscript.Echo "Копирует файл на диск A:. Там же создается копия файла."
    Wscript.Echo "Выполняет административные действия с базой 1С 8.2"
    Wscript.Echo ""
    Wscript.Echo "ПАРАМЕТРЫ ВЫЗОВА:"
    Wscript.Echo "  "+ WScript.ScriptName +" [файл-настроек | /? | /h]"
    Wscript.Echo ""
    Wscript.Echo "ПРИМЕР:"
    Wscript.Echo "1. cscript "+ WScript.ScriptName +" Файл.ini"
    Wscript.Echo "2. cscript "+ WScript.ScriptName
    Wscript.Echo "   Показывает этот экран."

End Sub

'********************************************************************
'* 
'* Function blnOpenFile
'*
'* Purpose: Opens a file.
'*
'* Input:   strFileName         A string with the name of the file.
'*
'* Output:  Sets objOpenFile to a FileSystemObject and setis it to 
'*            Nothing upon Failure.
'* 
'********************************************************************
Private Function blnOpenFile(ByVal strFileName, ByRef objOpenFile)

    ON ERROR RESUME NEXT

    If IsEmpty(strFileName) OR strFileName = "" Then
        blnOpenFile = False
        Set objOpenFile = Nothing
        Exit Function
    End If

    'fso.DeleteFile(strFileName)
    'Open the file for output
    Set objOpenFile = fso.CreateTextFile(strFileName, True)
    If blnErrorOccurred("Невозможно создать") Then
        blnOpenFile = False
        Set objOpenFile = Nothing
        Exit Function
    End If
    blnOpenFile = True

End Function

'********************************************************************
'*
'* Sub      VerifyHostIsCscript()
'*
'* Purpose: Determines which program is used to run this script.
'*
'* Input:   None
'*
'* Output:  If host is not cscript, then an error message is printed 
'*          and the script is aborted.
'*
'********************************************************************
Sub VerifyHostIsCscript()

    ON ERROR RESUME NEXT

    'Define constants
    CONST CONST_ERROR                   = 0
    CONST CONST_WSCRIPT                 = 1
    CONST CONST_CSCRIPT                 = 2
    
    Dim strFullName, strCommand, i, j, intStatus

    strFullName = WScript.FullName

    If Err.Number then
        Call Echo( "Error 0x" & CStr(Hex(Err.Number)) & " occurred." )
        If Err.Description <> "" Then
            Call Echo( "Error description: " & Err.Description & "." )
        End If
        intStatus =  CONST_ERROR
    End If

    i = InStr(1, strFullName, ".exe", 1)
    If i = 0 Then
        intStatus =  CONST_ERROR
    Else
        j = InStrRev(strFullName, "\", i, 1)
        If j = 0 Then
            intStatus =  CONST_ERROR
        Else
            strCommand = Mid(strFullName, j+1, i-j-1)
            Select Case LCase(strCommand)
                Case "cscript"
                    intStatus = CONST_CSCRIPT
                Case "wscript"
                    intStatus = CONST_WSCRIPT
                Case Else       'should never happen
                    Call Echo( "An unexpected program was used to " _
                                       & "run this script." )
                    Call Echo( "Only CScript.Exe or WScript.Exe can " _
                                       & "be used to run this script." )
                    intStatus = CONST_ERROR
                End Select
        End If
    End If

    If intStatus <> CONST_CSCRIPT Then
        Call Echo( "Please run this script using CScript." & vbCRLF & _
             "This can be achieved by" & vbCRLF & _
             "1. Using ""CScript SystemAccount.vbs arguments"" for Windows 95/98 or" _
             & vbCRLF & "2. Changing the default Windows Scripting Host " _
             & "setting to CScript" & vbCRLF & "    using ""CScript " _
             & "//H:CScript //S"" and running the script using" & vbCRLF & _
             "    ""SystemAccount.vbs arguments"" for Windows NT/2000/XP." )
        WScript.Quit(0)
    End If
End Sub 'VerifyHostIsCscript
