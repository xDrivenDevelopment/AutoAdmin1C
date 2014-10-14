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

Процедура ОбновитьИнформационнуюБазу()
	
	КоманднаяСтрокаРабочейБазы = мНастройки.ИмяСервера + "\" + мНастройки.ИмяБазы;
	
	ПараметрыСвязиСБазой = Новый Массив;
	ПараметрыСвязиСБазой.Добавить("ENTERPRISE");
	ПараметрыСвязиСБазой.Добавить("/S""" + КоманднаяСтрокаРабочейБазы + """");
	Если Не ПустаяСтрока(мНастройки.АдминистраторБазы) Тогда
		ПараметрыСвязиСБазой.Добавить("/N""" + мНастройки.АдминистраторБазы + """");
	КонецЕсли;
	Если Не ПустаяСтрока(мНастройки.ПарольАдминистратораБазы) Тогда
		ПараметрыСвязиСБазой.Добавить("/P""" + мНастройки.ПарольАдминистратораБазы + """");
	КонецЕсли;
	ПараметрыСвязиСБазой.Добавить("/CCLOSE");
	ПараметрыСвязиСБазой.Добавить("/UC""" + "AutoDeploy" + """");
	ПараметрыСвязиСБазой.Добавить("/DisableStartupMessages");
	ПараметрыСвязиСБазой.Добавить("/DisableStartupDialogs");
	ПараметрыСвязиСБазой.Добавить("/Out""" + ФайлИнформации() + """");
	
	СообщениеСборки("Обновление базы данных в режиме 1С:Предприятие");
	КодВозврата = ЗапуститьИПодождать(ПараметрыСвязиСБазой);
	Если КодВозврата = 0 Тогда
		СообщениеСборки("База данных обновлена: " + КоманднаяСтрокаРабочейБазы);
	Иначе
		СообщениеСборки("Не удалось обновить базу данных:");
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
ОбновитьИнформационнуюБазу();
