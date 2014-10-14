﻿
Перем мНастройки;
Перем мПараметрыДоступаКБазе;

Процедура ПрочитатьНастройки()
	
	Если АргументыКоманднойСтроки.Количество() = 0 Тогда
		Сообщить("Не заданы аргументы командной строки");
		Сообщить("Требуется задать:");
		Сообщить("<ПутьК1С>");
		ЗавершитьРаботу(1);
	КонецЕсли;
	
	СИ = Новый СистемнаяИнформация();
	Окружение = СИ.ПеременныеСреды();
	
	мНастройки = Новый Структура();
	мНастройки.Вставить("ПутьК1С", """" + АргументыКоманднойСтроки[0] + """");
	// Параметры сервера
	мНастройки.Вставить("ИмяСервера", Окружение["server_host"]);
	
	// Параметры рабочей базы
	мНастройки.Вставить("ИмяБазы", Окружение["db_name"]);
	мНастройки.Вставить("АдминистраторБазы", Окружение["db_user"]);
	мНастройки.Вставить("ПарольАдминистратораБазы", Окружение["db_password"]);
	
КонецПроцедуры

Процедура ОбновитьКонфигурациюБазыДанных()
	
	КоманднаяСтрокаРабочейБазы = мНастройки.ИмяСервера + "\" + мНастройки.ИмяБазы;
	
	ПараметрыСвязиСБазой = Новый Массив;
	ПараметрыСвязиСБазой.Добавить("DESIGNER");
	ПараметрыСвязиСБазой.Добавить("/S""" + КоманднаяСтрокаРабочейБазы + """");
	ПараметрыСвязиСБазой.Добавить("/N""" + мНастройки.АдминистраторБазы + """");
	ПараметрыСвязиСБазой.Добавить("/P""" + мНастройки.ПарольАдминистратораБазы + """");
	ПараметрыСвязиСБазой.Добавить("/UC""" + "AutoDeploy" + """");
	ПараметрыСвязиСБазой.Добавить("/DisableStartupMessages");
	ПараметрыСвязиСБазой.Добавить("/DisableStartupDialogs");
	ПараметрыСвязиСБазой.Добавить("/UpdateDBCfg -Server");
	ПараметрыСвязиСБазой.Добавить("/Out""" + ФайлИнформации() + """");
	
	СообщениеСборки("Обновление конфигурации базы данных");
	КодВозврата = ЗапуститьИПодождать(ПараметрыСвязиСБазой);
	Если КодВозврата = 0 Тогда
		СообщениеСборки("Конфигурация базы данных обновлена: " + КоманднаяСтрокаРабочейБазы);
	Иначе
		СообщениеСборки("Не удалось обновить конфигурацию базы данных:");
		ВывестиФайлИнформации();
		ЗавершитьРаботу(1);
	КонецЕсли;
	
КонецПроцедуры

Процедура СообщениеСборки(Знач Сообщение)

	Сообщить(Строка(ТекущаяДата()) + " " + Сообщение);
	
КонецПроцедуры

Функция ЗапуститьИПодождать(Параметры)

	СтрокаЗапуска = "";
	Для Каждого Параметр Из Параметры Цикл
	
		СтрокаЗапуска = СтрокаЗапуска + " " + Параметр;
	
	КонецЦикла;

	КодВозврата = 0;
	
	Сообщить(мНастройки.ПутьК1С + СтрокаЗапуска);
	
	ЗапуститьПриложение(мНастройки.ПутьК1С + СтрокаЗапуска, , Истина, КодВозврата);
	
	Возврат КодВозврата;

КонецФункции

Функция ФайлИнформации()
	Возврат "log.txt";
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

ПрочитатьНастройки();
ОбновитьКонфигурациюБазыДанных();
