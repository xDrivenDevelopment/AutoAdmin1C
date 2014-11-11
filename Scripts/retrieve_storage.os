﻿
Перем мМенеджерСборки;
Перем мНастройки;
Перем мКаталогСборки;
Перем мПараметрыДоступаКБазе;
Перем мПараметрыДоступаКХранилищу;

Процедура ПрочитатьНастройки()
	
	ПодключитьСценарий("Scripts/build_main.os", "МенеджерСборки");
	мМенеджерСборки = Новый МенеджерСборки();
	СообщениеСборки("Версия скрипта: " + мМенеджерСборки.ПолучитьВерсию());
	
	мНастройки = мМенеджерСборки.Настройки();
	
	ТекстПараметров = 
	" - Хранилище: <" + мНастройки.ПутьКХранилищу + ">
	| - Пользователь хранилища: <" + мНастройки.ПользовательХранилища + ">";
	
	СообщениеСборки("Параметры подключения к хранилищу:" + Символы.ПС + ТекстПараметров);
	
	ФайлСкрипта = Новый Файл(ТекущийСценарий().Источник);
	
	мКаталогСборки = ФайлСкрипта.Путь + "v8Temp";
	ОбеспечитьКаталог(мКаталогСборки);
	
	СообщениеСборки("Каталог сборки: " + мКаталогСборки);
	
	мПараметрыДоступаКБазе = Новый Массив;
	мПараметрыДоступаКБазе.Добавить("DESIGNER");
	мПараметрыДоступаКБазе.Добавить("/F""" + ПутьКВременнойБазе() + """");
	мПараметрыДоступаКБазе.Добавить("/Out""" + ФайлИнформации() + """");
	
	мПараметрыДоступаКХранилищу = СкопироватьМассив(мПараметрыДоступаКБазе);
	мПараметрыДоступаКХранилищу.Добавить("/DisableStartupMessages");
	мПараметрыДоступаКХранилищу.Добавить("/DisableStartupDialogs");
	мПараметрыДоступаКХранилищу.Добавить("/ConfigurationRepositoryF """+мНастройки.ПутьКХранилищу+"""");
	мПараметрыДоступаКХранилищу.Добавить("/ConfigurationRepositoryN """+мНастройки.ПользовательХранилища+"""");
	Если Не ПустаяСтрока(мНастройки.ПарольХранилища) Тогда
		мПараметрыДоступаКХранилищу.Добавить("/ConfigurationRepositoryP """+мНастройки.ПарольХранилища+"""");
	КонецЕсли;
	
КонецПроцедуры

Процедура ОбеспечитьКаталог(Знач Каталог)

	Файл = Новый Файл(Каталог);
	Если НЕ Файл.Существует() Тогда
		СоздатьКаталог(Каталог);
	ИначеЕсли Не Файл.ЭтоКаталог() Тогда
		ВызватьИсключение "Каталог " + Каталог + " не является каталогом";
	КонецЕсли;

КонецПроцедуры

Функция СкопироватьМассив(Источник)
	
	НовыйМассив = Новый Массив;
	Для Каждого Элемент Из Источник Цикл
		НовыйМассив.Добавить(Элемент);
	КонецЦикла;
	
	Возврат НовыйМассив;
	
КонецФункции

Функция ЗапуститьИПодождать(Параметры)

	Возврат мМенеджерСборки.ЗапуститьИПодождать(Параметры);

КонецФункции

Функция ФайлИнформации()
	Возврат мКаталогСборки + "\log.txt";
КонецФункции

Функция ПутьКВременнойБазе()
	Возврат мКаталогСборки + "\TempDB";
КонецФункции

Процедура ВывестиФайлИнформации()

	Файл = Новый Файл(ФайлИнформации());
	Если Файл.Существует() Тогда
		Чтение = Новый ЧтениеТекста(Файл.ПолноеИмя);
		Сообщение = Чтение.Прочитать();
		Сообщить(Сообщение);
		Чтение.Закрыть();
	Иначе
		Сообщить("Информации об ошибке нет");
	КонецЕсли;

КонецПроцедуры

Процедура СообщениеСборки(Знач Сообщение)

	Сообщить(Строка(ТекущаяДата()) + " " + Сообщение);
	
КонецПроцедуры

Функция ПараметрыВременнойБазы()

	ПараметрыВремБазы = СкопироватьМассив(мПараметрыДоступаКБазе);
	ПараметрыВремБазы[1] = "/F""" + ПутьКВременнойБазе() + """";
	ПараметрыВремБазы.Удалить(2);
	
	Возврат ПараметрыВремБазы;

КонецФункции

Процедура СоздатьВременнуюБазу()

	СообщениеСборки("Создание временной базы для запуска Конфигуратора");
	КаталогВременнойБазы = ПутьКВременнойБазе();
	ОбеспечитьКаталог(КаталогВременнойБазы);
	УдалитьФайлы(КаталогВременнойБазы, "*.*");
	
	ПараметрыЗапуска = Новый Массив;
	ПараметрыЗапуска.Добавить("CREATEINFOBASE");
	ПараметрыЗапуска.Добавить("File="""+КаталогВременнойБазы+""";");
	ПараметрыЗапуска.Добавить("/Out""" + ФайлИнформации() + """");
	
	КодВозврата = ЗапуститьИПодождать(ПараметрыЗапуска);
	Если КодВозврата = 0 Тогда
		СообщениеСборки("Временная база создана");
	Иначе
		СообщениеСборки("Не удалось создать временную базу:");
		ВывестиФайлИнформации();
		ЗавершитьРаботу(1);
	КонецЕсли;

КонецПроцедуры

Процедура ПолучитьВерсиюИзХранилища()

	ПараметрыПолучитьВерсию = СкопироватьМассив(мПараметрыДоступаКХранилищу);
	ПараметрыПолучитьВерсию.Добавить("/ConfigurationRepositoryDumpCfg """+мКаталогСборки+"\source.cf""");

	СообщениеСборки("Получение версии из хранилища");
	КодВозврата = ЗапуститьИПодождать(ПараметрыПолучитьВерсию);
	Если КодВозврата = 0 Тогда
		ВывестиФайлИнформации();
		СообщениеСборки("Версия получена");
	Иначе
		СообщениеСборки("Не удалось получить версию из хранилища:");
		ВывестиФайлИнформации();
		ЗавершитьРаботу(1);
	КонецЕсли;
	
КонецПроцедуры

Процедура ПроверитьПоявлениеВерсии()
	
	СообщениеСборки("Проверка существования файла версии");
	
	ФайлВерсии = Новый Файл(мКаталогСборки + "\source.cf");
	Если Не ФайлВерсии.Существует() Тогда
		ВызватьИсключение "Файл с актуальной версией не обнаружен";
	КонецЕсли;
	
	СообщениеСборки("Файл версии создан");
	
КонецПроцедуры

Процедура ОбновитьКонфигурациюИзХранилища()

	КоманднаяСтрокаРабочейБазы = мНастройки.ИмяСервера + "\" + мНастройки.ИмяБазы;
	
	ПараметрыСвязиСБазой = Новый Массив;
	ПараметрыСвязиСБазой.Добавить("DESIGNER");
	ПараметрыСвязиСБазой.Добавить("/S""" + КоманднаяСтрокаРабочейБазы + """");
	Если Не ПустаяСтрока(мНастройки.АдминистраторБазы) Тогда
		ПараметрыСвязиСБазой.Добавить("/N""" + мНастройки.АдминистраторБазы + """");
	КонецЕсли;
	Если Не ПустаяСтрока(мНастройки.ПарольАдминистратораБазы) Тогда
		ПараметрыСвязиСБазой.Добавить("/P""" + мНастройки.ПарольАдминистратораБазы + """");
	КонецЕсли;
	ПараметрыСвязиСБазой.Добавить("/WA+");
	ПараметрыСвязиСБазой.Добавить("/UC""" + мНастройки.КодРазрешения + """");
	
	Для Сч = 2 По мПараметрыДоступаКХранилищу.Количество() - 1 Цикл
		Параметр = мПараметрыДоступаКХранилищу[Сч];
		ПараметрыСвязиСБазой.Добавить(Параметр);
	КонецЦикла;
	
	ПараметрыСвязиСБазой.Добавить("/ConfigurationRepositoryUpdateCfg -force -revised");
	
	СообщениеСборки("Обновление конфигурации из хранилища");
	КодВозврата = ЗапуститьИПодождать(ПараметрыСвязиСБазой);
	Если КодВозврата = 0 Тогда
		ВывестиФайлИнформации();
		СообщениеСборки("Конфигурация обновлена");
	Иначе
		СообщениеСборки("Не удалось обновить конфигурацию из хранилища:");
		ВывестиФайлИнформации();
		ЗавершитьРаботу(1);
	КонецЕсли;
	
КонецПроцедуры

Процедура ВыгрузитьВерсиюКонфигурацииИзХранилища()

	СоздатьВременнуюБазу();
	ПолучитьВерсиюИзХранилища();
	ПроверитьПоявлениеВерсии();

КонецПроцедуры

ПрочитатьНастройки();

Если АргументыКоманднойСтроки.Количество() = 0 Тогда
	ОбновитьКонфигурациюИзХранилища();	
ИначеЕсли АргументыКоманднойСтроки[0] = "-dumpstorage" Тогда
	ВыгрузитьВерсиюКонфигурацииИзХранилища();
Иначе
	СообщениеСборки("Неизвестный параметр командной строки <"+АргументыКоманднойСтроки[0]+">");
	ЗавершитьРаботу(1);
КонецЕсли;




