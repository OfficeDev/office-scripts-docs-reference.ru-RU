---
title: Справочник API для сценариев Office
description: Общие сведения об API JavaScript для сценариев Office
ms.date: 12/12/2019
ms.openlocfilehash: b3e75cbecac13f41c83f019c681b0682571ead41
ms.sourcegitcommit: 2834ee43a14a4b0953d5fb22749c115a42960e20
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/10/2020
ms.locfileid: "42705947"
---
# <a name="office-scripts-api-reference"></a>Справочник API для сценариев Office

API сценариев Office позволяет автоматизировать выполнение распространенных задач в Excel в Интернете. Используйте эту справочную документацию, чтобы узнать больше о классах, методах и других типах, доступных для ваших сценариев. Все объекты, доступные с помощью скриптов Office, можно найти в оглавлении слева от страницы.

## <a name="common-classes"></a>Общие классы

Приведенный ниже список разделяет основные принципы объектной модели сценариев Office. Здесь показаны общие классы и их связь друг с другом.

- [Книга](/javascript/api/office-scripts/excel/excel.workbook) содержит один или несколько [листов](/javascript/api/office-scripts/excel/excel.worksheet) в [WorksheetCollection](/javascript/api/office-scripts/excel/excel.worksheetcollection).
- [Рабочий лист](/javascript/api/office-scripts/excel/excel.worksheet) предоставляет доступ к ячейкам через объекты [Range](/javascript/api/office-scripts/excel/excel.range).
- [Range](/javascript/api/office-scripts/excel/excel.range) представляет группу смежных клеток.
- [Диапазоны](/javascript/api/office-scripts/excel/excel.range) используются для создания и размещения [таблиц](/javascript/api/office-scripts/excel/excel.table), [диаграмм](/javascript/api/office-scripts/excel/excel.chart), [фигур](/javascript/api/office-scripts/excel/excel.shape) и других объектов визуализации данных или организации.
- [Лист](/javascript/api/office-scripts/excel/excel.worksheet) содержит коллекции этих объектов данных (например, [ChartCollection](/javascript/api/office-scripts/excel/excel.chartcollection)), присутствующих на отдельном листе.
- [Книги](/javascript/api/office-scripts/excel/excel.workbook). содержат коллекции некоторых объектов данных (например, [TableCollection](/javascript/api/office-scripts/excel/excel.tablecollection)) для всей [книги](/javascript/api/office-scripts/excel/excel.workbook).

Дополнительные сведения об объектной модели сценариев Office можно найти [в разделе Основные сведения о сценариях для сценариев Office в Excel в Интернете](/office/dev/scripts/develop/scripting-fundamentals)

## <a name="see-also"></a>См. также

- [Сведения о сценариях Office](/office/dev/scripts/overview/excel)
- [Запись, редактирование и создание сценариев Office в Excel в Интернете](/office/dev/scripts/tutorials/excel-tutorial)
- [Основы сценариев для сценариев Office в Excel в Интернете](/office/dev/scripts/develop/scripting-fundamentals)
