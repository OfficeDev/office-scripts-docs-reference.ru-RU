---
title: Справочник API для сценариев Office
description: Обзор сценариев Office для асинхронных API JavaScript.
ms.date: 06/17/2020
ms.openlocfilehash: b9fa59764b7cb54567adbe05b9671bae9b232972
ms.sourcegitcommit: 163b26a43411ad7f13a01237efe9b8d6de656b47
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 06/25/2020
ms.locfileid: "44887085"
---
# <a name="office-scripts-async-api-reference"></a>Справочник по асинхронным API сценариев Office

Асинхронный API сценариев Office поддерживает старые сценарии, созданные на этапе предварительной версии сценариев Office. Используйте эту справочную документацию, чтобы узнать больше о классах, методах и других типах, используемых этими старыми сценариями. Все объекты, доступные с помощью скриптов Office, можно найти в оглавлении слева от страницы.

> [!IMPORTANT]
> Мы настоятельно рекомендуем создавать новые сценарии с помощью стандартных API сценариев Office. При создании или редактировании нового скрипта переключитесь на [неасинхронную версию](?view=office-scripts) API.

## <a name="common-classes"></a>Общие классы

Приведенный ниже список разделяет основные принципы объектной модели сценариев Office. Здесь показаны общие классы и их связь друг с другом.

- [Книга](/javascript/api/office-scripts/excel/excelscript.workbook) содержит один или несколько [листов](/javascript/api/office-scripts/excel/excelscript.worksheet) в [WorksheetCollection](/javascript/api/office-scripts/excel/excelscript.worksheetcollection).
- [Рабочий лист](/javascript/api/office-scripts/excel/excelscript.worksheet) предоставляет доступ к ячейкам через объекты [Range](/javascript/api/office-scripts/excel/excelscript.range).
- [Range](/javascript/api/office-scripts/excel/excelscript.range) представляет группу смежных клеток.
- [Диапазоны](/javascript/api/office-scripts/excel/excelscript.range) используются для создания и размещения [таблиц](/javascript/api/office-scripts/excel/excelscript.table), [диаграмм](/javascript/api/office-scripts/excel/excelscript.chart), [фигур](/javascript/api/office-scripts/excel/excelscript.shape) и других объектов визуализации данных или организации.
- [Лист](/javascript/api/office-scripts/excel/excelscript.worksheet) содержит коллекции этих объектов данных (например, [ChartCollection](/javascript/api/office-scripts/excel/excelscript.chartcollection)), присутствующих на отдельном листе.
- [Книги](/javascript/api/office-scripts/excel/excelscript.workbook) содержат коллекции некоторых объектов данных (например, [TableCollection](/javascript/api/office-scripts/excel/excelscript.tablecollection)) для всей [книги](/javascript/api/office-scripts/excel/excelscript.workbook).

Дополнительные сведения об объектной модели сценариев Office можно найти [в разделе Основные сведения о сценариях для сценариев Office в Excel в Интернете](/office/dev/scripts/develop/scripting-fundamentals)

## <a name="see-also"></a>См. также

- [Использование асинхронных API сценариев Office для поддержки устаревших скриптов](/office/dev/scripts/develop/excel-async-model)
- [Сведения о сценариях Office](/office/dev/scripts/overview/excel)
- [Запись, редактирование и создание сценариев Office в Excel в Интернете](/office/dev/scripts/tutorials/excel-tutorial)
- [Основы сценариев для сценариев Office в Excel в Интернете](/office/dev/scripts/develop/scripting-fundamentals)
