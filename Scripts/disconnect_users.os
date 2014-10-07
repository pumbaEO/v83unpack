﻿
Перем мНастройки; // соответствие настроек скрипта
Перем мПредыдущиеСвойстваИБ; // Состояния флагов блокировки ИБ до начала выполнения скрипта

////////////////////////////////////////////////////////////////////////////
// Инициализация скрипта

Процедура ПрочитатьНастройки()
	
	мНастройки = Новый Структура;
	СИ = Новый СистемнаяИнформация();
	
	Окружение = СИ.ПеременныеСреды();
	
	// Параметры сервера
	мНастройки.Вставить("ИмяСервера", Окружение["server_host"]);
	мНастройки.Вставить("АдминистраторКластера", Окружение["cluster_admin"]);
	мНастройки.Вставить("ПарольАдминистратораКластера", Окружение["cluster_admin_password"]);
	мНастройки.Вставить("КлассCOMСоединения", Окружение["com_connector"]);
	мНастройки.Вставить("АдминистраторКластера", Окружение["cluster_admin"]);
	
	// Параметры рабочей базы
	мНастройки.Вставить("ИмяБазы", Окружение["db_name"]);
	мНастройки.Вставить("АдминистраторБазы", Окружение["db_user"]);
	мНастройки.Вставить("ПарольАдминистратораБазы", Окружение["db_password"]);
	
	// Прочие настройки
	мНастройки.Вставить("СообщениеБлокировки", Окружение["lock_message"]);
	мНастройки.Вставить("ТаймаутБлокировки", Окружение["lock_timeout"]);
	
	Если мНастройки.ТаймаутБлокировки = Неопределено Тогда
		мНастройки.ТаймаутБлокировки = 1000;
	КонецЕсли;
	
КонецПроцедуры

Процедура ПроверитьОбязательныеНастройки()
	
	Если ПустаяСтрока(мНастройки.КлассCOMСоединения) Тогда
		ВызватьИсключение "Не задан класс COM-соединения";
	КонецЕсли;
	
	Если ПустаяСтрока(мНастройки.ИмяСервера) Тогда
		ВызватьИсключение "Не задано имя сервера приложений 1С";
	КонецЕсли;
	
	Если ПустаяСтрока(мНастройки.ИмяБазы) Тогда
		ВызватьИсключение "Не задано имя базы данных 1С";
	КонецЕсли;
	
КонецПроцедуры


////////////////////////////////////////////////////////////////////////////
// Основная полезная нагрузка

Функция ОтключитьПользователей()
	
	Перем ComConnector;
	Перем ServerAgent;
	Перем Clusters;
	
	ИмяСервера = мНастройки.ИмяСервера;
	ComConnector = ПолучитьСоединениеСКластером();

    СообщениеСборки("Подключение к агенту сервера");
	ServerAgent = ComConnector.ConnectAgent(ИмяСервера);

    СообщениеСборки("Получение массива кластеров сервера у агента сервера");
    Clusters = ServerAgent.GetClusters();

    СообщениеСборки("Начало завершения работы пользователей");
	
	СоединенияОтключены = Истина;
	Cluster = Неопределено;
	Попытка
		
		Cluster = НайтиКластерСерверов(Clusters, ИмяСервера);
		
		СообщениеСборки("Аутентикация к найденному кластеру: " + Cluster.ClusterName + ", "+Cluster.HostName);
		ServerAgent.Authenticate(Cluster, мНастройки.АдминистраторКластера, мНастройки.ПарольАдминистратораКластера);
		
		СообщениеСборки("Получение списка работающих рабочих процессов и обход в цикле");
		ОтключитьСоединенияПоРабочимСерверам(ComConnector, ServerAgent, Cluster);
		
		СоединенияОтключены = Истина;
		
	Исключение
		СоединенияОтключены = Ложь;
		Сообщить(ИнформацияОбОшибке().Описание);
	КонецПопытки;
	
	ОсвободитьОбъектКластера(Cluster);
	ОсвободитьОбъектКластера(Clusters);
	ОсвободитьОбъектКластера(ServerAgent);
	ОсвободитьОбъектКластера(ComConnector);
	
	Возврат СоединенияОтключены;
	
КонецФункции

Функция ПолучитьСоединениеСКластером()
	
	Соединение = мНастройки.КлассCOMСоединения;
	СообщениеСборки("Создание COM-коннектора <"+ Соединение + ">");
	
	Возврат Новый COMОбъект(Соединение);
	
КонецФункции

Функция НайтиКластерСерверов(Знач Clusters, Знач ИмяСервера)
	
	НашлиКластер = Ложь;
	Для i = 0 По Clusters.Количество()-1 Цикл
		Cluster = Clusters[i];
		Если ВРег(Cluster.HostName) = ВРег(ИмяСервера) Тогда
			НашлиКластер = Истина;
			Прервать;
		КонецЕсли;
		
		ОсвободитьОбъектКластера(Cluster);
		
	КонецЦикла;
	
	Если Не НашлиКластер Тогда
		ОсвободитьОбъектКластера(Cluster);
		ВызватьИсключение "Ошибка - не нашли кластер <"+ИмяСервера+">";
	КонецЕсли;
	
	Возврат Cluster;
	
КонецФункции

Процедура ОтключитьСоединенияПоРабочимСерверам(Знач ComConnector, Знач ServerAgent, Знач Cluster)
	
	Попытка
		WorkServers = ServerAgent.GetWorkingServers(Cluster);
	    Для i = 0 По WorkServers.Количество()-1 Цикл
	        WorkServer = WorkServers[i];
			
	        СообщениеСборки("Обрабатываю рабочий сервер "+WorkServer.Name+": " + WorkServer.HostName);

	        WorkingProcesses = ServerAgent.GetServerWorkingProcesses(Cluster, WorkServer);

			НашлиИнформационнуюБазу = Ложь;
			
	        Для j = 0 To WorkingProcesses.Количество()-1 Цикл

				Если WorkingProcesses[j].Running = 1 Тогда
					
					СтрокаСоединения = "tcp://" + WorkingProcesses[j].HostName + ":" + WorkingProcesses[j].MainPort;
					СообщениеСборки("Создание соединения с рабочим процессом " + СтрокаСоединения);
	                ConnectToWorkProcess = ComConnector.ConnectWorkingProcess(СтрокаСоединения);
					
					ConnectToWorkProcess.AuthenticateAdmin(мНастройки.АдминистраторКластера, мНастройки.ПарольАдминистратораКластера);
	                ConnectToWorkProcess.AddAuthentication(мНастройки.АдминистраторБазы, мНастройки.ПарольАдминистратораБазы);
					
					СвойстваБлокировки = ОтключитьСоединенияВРабочемПроцессе(ServerAgent, Cluster, ConnectToWorkProcess);
					Если мПредыдущиеСвойстваИБ = Неопределено Тогда
						мПредыдущиеСвойстваИБ = СвойстваБлокировки;
					КонецЕсли;
					
					ОсвободитьОбъектКластера(ConnectToWorkProcess);
					
				КонецЕсли;
				
			КонецЦикла;
			
			ОсвободитьОбъектКластера(WorkingProcesses);
			ОсвободитьОбъектКластера(WorkServer);
			
		КонецЦикла;
		
	Исключение
		ОсвободитьОбъектКластера(ConnectToWorkProcess);
		ОсвободитьОбъектКластера(WorkingProcesses);
		ОсвободитьОбъектКластера(WorkServer);
		ОсвободитьОбъектКластера(WorkServers);
		ВызватьИсключение;
	КонецПопытки;
	
	ОсвободитьОбъектКластера(ConnectToWorkProcess);
	ОсвободитьОбъектКластера(WorkingProcesses);
	ОсвободитьОбъектКластера(WorkServer);
	ОсвободитьОбъектКластера(WorkServers);
	
КонецПроцедуры

Функция ОтключитьСоединенияВРабочемПроцессе(Знач ServerAgent, Знач Cluster, Знач ConnectToWorkProcess)
	
	ПредыдущиеБлокировки = ЗаблокироватьСоединенияИБ(ConnectToWorkProcess);
	
	СообщениеСборки("Задержка перед началом завершения работы пользователей "+Цел(мНастройки.ТаймаутБлокировки/1000) + " секунд");
	
    InfoBaseSession = НайтиСеансИнформационнойБазы(ServerAgent, Cluster);
	Если InfoBaseSession = Неопределено Тогда
		ВызватьИсключение "Не нашли нужную ИБ для сессии";
	КонецЕсли;
	
	СообщениеСборки("Начало завершение работы пользователей с ИБ " + мНастройки.ИмяБазы);
	Попытка
		УдалитьСеансыИнформационнойБазы(ServerAgent, Cluster, InfoBaseSession);	
	Исключение
		ВосстановитьСостояниеБлокировкиИБ(ConnectToWorkProcess, Неопределено, ПредыдущиеБлокировки.СтатусБлокировки, ПредыдущиеБлокировки.СтатусРегЗаданий);
		ОсвободитьОбъектКластера(InfoBaseSession);
		ВызватьИсключение;
	КонецПопытки;
	
	ОсвободитьОбъектКластера(InfoBaseSession);
	СообщениеСборки("Окончание завершения работы пользователей:" + мНастройки.ИмяБазы);	
	
	Возврат ПредыдущиеБлокировки;
	
КонецФункции

Функция ЗаблокироватьСоединенияИБ(Знач ConnectToWorkProcess)
	
	InfoBase = НайтиИнформационнуюБазуВРабочемПроцессе(ConnectToWorkProcess);
	Если Infobase = Неопределено Тогда
		ВызватьИсключение "Не нашли нужную ИБ";
	КонецЕсли;
	
	ТекущийСтатусРегЗаданий = InfoBase.ScheduledJobsDenied;
	ТекущийСтатусБлокировки = InfoBase.ConnectDenied;
	
	СообщениеСборки("Установка запрета на подключения к ИБ: " + InfoBase.Name);
	
    InfoBase.ConnectDenied = Истина;
    InfoBase.ScheduledJobsDenied = Истина;
    InfoBase.DeniedMessage = мНастройки.СообщениеБлокировки;
    InfoBase.PermissionCode = мНастройки.ПарольАдминистратораБазы;
	
	Попытка
		ConnectToWorkProcess.UpdateInfoBase(InfoBase);
	Исключение
		
		ТекстОшибки = ИнформацияОбОшибке().Описание;
		СообщениеСборки("Не удалось заблокировать подключения: <" + ТекстОшибки + "> Попытка восстановления...");
		
		Попытка
			ВосстановитьСостояниеБлокировкиИБ(ConnectToWorkProcess, InfoBase, ТекущийСтатусБлокировки, ТекущийСтатусРегЗаданий);
		Исключение
			ОсвободитьОбъектКластера(InfoBase);
			ВызватьИсключение;	
		КонецПопытки;
		
		ОсвободитьОбъектКластера(InfoBase);
		ВызватьИсключение;
		
	КонецПопытки;
	
	ОсвободитьОбъектКластера(InfoBase);
	
	ОпцииБлокировки = Новый Структура();
	ОпцииБлокировки.Вставить("СтатусБлокировки", ТекущийСтатусБлокировки);
	ОпцииБлокировки.Вставить("СтатусРегЗаданий", ТекущийСтатусРегЗаданий);
	
	Возврат ОпцииБлокировки;
	
КонецФункции

Функция НайтиСеансИнформационнойБазы(Знач ServerAgent, Знач Cluster)
	
	СообщениеСборки("Поиск нужной ИБ для сессии");
	Возврат НайтиИнформационнуюБазуВКоллекции(ServerAgent.GetInfoBases(Cluster));
	
КонецФункции

Процедура УдалитьСеансыИнформационнойБазы(Знач ServerAgent, Знач Cluster, Знач InfoBaseSession)
	
	СообщениеСборки("Обработка списка сеансов");
	
	Sessions = ServerAgent.GetInfoBaseSessions(Cluster, InfoBaseSession);
    Для Сч = 0 По Sessions.Количество()-1 Цикл
        Session  = Sessions[Сч];
        UserName = Session.UserName;
        AppID    = ВРег(Session.AppID);
        
        СообщениеСборки("Попытка отключения: " + "User=["+UserName+"] ConnID=["+""+"] AppID=["+AppID+"]");
        ServerAgent.TerminateSession(Cluster, Session);
		ОсвободитьОбъектКластера(Session);
		СообщениеСборки("Выполнено");
	КонецЦикла;
	
КонецПроцедуры

Функция НайтиИнформационнуюБазуВРабочемПроцессе(Знач ConnectToWorkProcess)
	
	СообщениеСборки("Получение списка ИБ рабочего процесса");
	Возврат НайтиИнформационнуюБазуВКоллекции(ConnectToWorkProcess.GetInfoBases());
	
КонецФункции

Процедура ВосстановитьСостояниеБлокировкиИБ(ConnectToWorkProcess, InfoBase, СтатусБлокировки, СтатусРегЗаданий)
	
	Если InfoBase = Неопределено Тогда
		InfoBase = НайтиИнформационнуюБазуВРабочемПроцессе(ConnectToWorkProcess);
		Если Infobase = Неопределено Тогда
			ВызватьИсключение "Не нашли нужную ИБ при попытке восстановления блокировки";
		КонецЕсли;
	КонецЕсли;
	
	Попытка
		InfoBase.ConnectDenied = СтатусБлокировки;
		InfoBase.ScheduledJobsDenied = СтатусРегЗаданий;
		InfoBase.DeniedMessage = "";
		InfoBase.PermissionCode = "";
		ConnectToWorkProcess.UpdateInfoBase(InfoBase);
	Исключение
		СообщениеСборки("Не удалось восстановить опции блокировки:" + ИнформацияОбОшибке().Описание);
		ВызватьИсключение;
	КонецПопытки;
	
КонецПроцедуры

Функция НайтиИнформационнуюБазуВКоллекции(Знач InfoBases)
	
	Перем InfoBase;
	
	Попытка
		ИскомаяИБ = мНастройки.ИмяБазы;
		БазаНайдена = Ложь;
		
		InfoBase = ОбойтиКоллекциюИНайтиИБ(InfoBases, ИскомаяИБ);
		
		БазаНайдена = InfoBase <> Неопределено;
		
	Исключение
		ОсвободитьОбъектКластера(InfoBase);
		ОсвободитьОбъектКластера(InfoBases);
		ВызватьИсключение;
	КонецПопытки;
	
	Если Не БазаНайдена Тогда
		InfoBase = Неопределено;
	КонецЕсли;
	
	ОсвободитьОбъектКластера(InfoBases);
	
	Возврат InfoBase;
	
КонецФункции

Функция ОбойтиКоллекциюИНайтиИБ(Знач InfoBases,Знач ИскомаяИБ)
	
	Перем InfoBase;
	СообщениеСборки("Поиск ИБ " + ИскомаяИБ);
    Для Каждого InfoBase Из InfoBases Цикл
        СообщениеСборки(" Обрабатывается ИБ: " + InfoBase.Name);
        Если НРег(InfoBase.Name) = НРег(ИскомаяИБ) Then
            БазаНайдена = Истина;
            СообщениеСборки(" Нашли нужную ИБ");
            Прервать;
		КонецЕсли;
	КонецЦикла;
	
	Если Не БазаНайдена Тогда
		ОсвободитьОбъектКластера(InfoBase);
	КонецЕсли;
	
	Возврат InfoBase;
	
КонецФункции


////////////////////////////////////////////////////////////////////////////
// Служебные процедуры

Процедура СообщениеСборки(Знач Сообщение)

	Сообщить(Строка(ТекущаяДата()) + " " + Сообщение);
	
КонецПроцедуры

Процедура ОсвободитьОбъектКластера(Соединение)
	
	Если Соединение <> Неопределено Тогда
		ОсвободитьОбъект(Соединение);
		Соединение = Неопределено;
	КонецЕсли;
	
КонецПроцедуры


////////////////////////////////////////////////////////////////////////////
// Точка входа в скрипт

ПрочитатьНастройки();
ПроверитьОбязательныеНастройки();

Если Не ОтключитьПользователей() Тогда
	СообщениеСборки("Отключение не выполнено. См. журнал сообщений");
	ЗавершитьРаботу(1);
КонецЕсли;