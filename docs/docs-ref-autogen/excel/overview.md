---
title: Справочник API для сценариев Office
description: Обзор API Office скриптов JavaScript.
ms.date: 06/29/2020
ms.openlocfilehash: b9ac5b191dbbb72301fc1f4fe11c81bd2b45ae17
ms.sourcegitcommit: dfc12de9e6eb5de71199b36f92ce93039509ad37
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/23/2022
ms.locfileid: "63755993"
---
# <a name="office-scripts-api-reference"></a>Справочник API для сценариев Office

API Office скриптов позволяет автоматизировать общие задачи в Excel в Интернете. Используйте эту справочную документацию, чтобы узнать больше о классах, методах и других типах, доступных для сценариев. Все объекты, доступные Office скрипты, можно найти в таблице содержимого слева от страницы.

> [!NOTE]
> Если вы ищете API JavaScript для разработки Office надстройки, посетите ссылку Office [API JavaScript](/javascript/api/overview?view=excel-js-preview&preserve-view=true).

## <a name="common-classes"></a>Общие классы

Следующий список разбивает основы объектной модели Office Scripts. Это показывает общие классы и их отношение друг к другу.

- [Рабочая книга](/javascript/api/office-scripts/excelscript/excelscript.workbook) содержит одну или несколько [рабочих листов](/javascript/api/office-scripts/excelscript/excelscript.worksheet).
- [Рабочий лист](/javascript/api/office-scripts/excelscript/excelscript.worksheet) предоставляет доступ к ячейкам через объекты [Range](/javascript/api/office-scripts/excelscript/excelscript.range).
- [Range](/javascript/api/office-scripts/excelscript/excelscript.range) представляет группу смежных клеток.
- [Диапазоны](/javascript/api/office-scripts/excelscript/excelscript.range) используются для создания и размещения [таблиц](/javascript/api/office-scripts/excelscript/excelscript.table), [диаграмм](/javascript/api/office-scripts/excelscript/excelscript.chart), [фигур](/javascript/api/office-scripts/excelscript/excelscript.shape) и других объектов визуализации данных или организации.
- Лист [содержит массивы](/javascript/api/office-scripts/excelscript/excelscript.worksheet) , заполненные теми объектами, которые присутствуют в отдельном листе.
- Книга [содержит](/javascript/api/office-scripts/excelscript/excelscript.workbook) массивы некоторых из этих объектов данных для всей книги.

Дополнительные сведения об объектной модели Office Scripts можно получить в Office [scripts in Excel в Интернете](/office/dev/scripts/develop/scripting-fundamentals)

## <a name="see-also"></a>См. также

- [О Office скриптах](/office/dev/scripts/overview/excel)
- [Запись, редактирование и создание сценариев Office в Excel в Интернете](/office/dev/scripts/tutorials/excel-tutorial)
- [Основные сведения о сценариях Office в Excel для Интернета](/office/dev/scripts/develop/scripting-fundamentals)
