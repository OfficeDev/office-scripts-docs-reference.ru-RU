---
title: Справочник API для сценариев Office
description: Общие сведения о сценариях API JavaScript для Office.
ms.date: 06/17/2020
ms.openlocfilehash: 5634d0e5f68464655054ad1c09eb7931e0da62d4
ms.sourcegitcommit: 163b26a43411ad7f13a01237efe9b8d6de656b47
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 06/25/2020
ms.locfileid: "44884837"
---
# <a name="office-scripts-api-reference"></a>Справочник API для сценариев Office

API сценариев Office позволяет автоматизировать выполнение распространенных задач в Excel в Интернете. Используйте эту справочную документацию, чтобы узнать больше о классах, методах и других типах, доступных для ваших сценариев. Все объекты, доступные с помощью скриптов Office, можно найти в оглавлении слева от страницы.

> [!NOTE]
> Если вы ищете API JavaScript для разработки надстроек Office, посетите [Справочник по API JavaScript](/javascript/api/overview?view=excel-js-preview)для надстроек Office.

## <a name="common-classes"></a>Общие классы

Приведенный ниже список разделяет основные принципы объектной модели сценариев Office. Здесь показаны общие классы и их связь друг с другом.

- [Рабочая книга](/javascript/api/office-scripts/excel/excelscript.workbook) содержит одну или несколько [рабочих листов](/javascript/api/office-scripts/excel/excelscript.worksheet).
- [Рабочий лист](/javascript/api/office-scripts/excel/excelscript.worksheet) предоставляет доступ к ячейкам через объекты [Range](/javascript/api/office-scripts/excel/excelscript.range).
- [Range](/javascript/api/office-scripts/excel/excelscript.range) представляет группу смежных клеток.
- [Диапазоны](/javascript/api/office-scripts/excel/excelscript.range) используются для создания и размещения [таблиц](/javascript/api/office-scripts/excel/excelscript.table), [диаграмм](/javascript/api/office-scripts/excel/excelscript.chart), [фигур](/javascript/api/office-scripts/excel/excelscript.shape) и других объектов визуализации данных или организации.
- [Лист](/javascript/api/office-scripts/excel/excelscript.worksheet) содержит массивы, заполненные объектами, присутствующими на отдельном листе.
- [Книга](/javascript/api/office-scripts/excel/excelscript.workbook) содержит массивы некоторых из этих объектов данных для всей книги.

Дополнительные сведения об объектной модели сценариев Office можно найти [в разделе Основные сведения о сценариях для сценариев Office в Excel в Интернете](/office/dev/scripts/develop/scripting-fundamentals)

## <a name="see-also"></a>См. также

- [Сведения о сценариях Office](/office/dev/scripts/overview/excel)
- [Запись, редактирование и создание сценариев Office в Excel в Интернете](/office/dev/scripts/tutorials/excel-tutorial)
- [Основы сценариев для сценариев Office в Excel в Интернете](/office/dev/scripts/develop/scripting-fundamentals)
