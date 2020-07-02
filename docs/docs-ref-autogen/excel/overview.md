---
title: Справочник API для сценариев Office
description: Общие сведения о сценариях API JavaScript для Office.
ms.date: 06/29/2020
ms.openlocfilehash: 7c4fe97ca35cfb442ebbf9db2e0b03b389185ae8
ms.sourcegitcommit: 9c4c4c213a203e58c55eb3d84d7d92fa527f3eb8
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/01/2020
ms.locfileid: "45004733"
---
# <a name="office-scripts-api-reference"></a>Справочник API для сценариев Office

API сценариев Office позволяет автоматизировать выполнение распространенных задач в Excel в Интернете. Используйте эту справочную документацию, чтобы узнать больше о классах, методах и других типах, доступных для ваших сценариев. Все объекты, доступные с помощью скриптов Office, можно найти в оглавлении слева от страницы.

> [!NOTE]
> Если вы ищете API JavaScript для разработки надстроек Office, посетите [Справочник по API JavaScript](/javascript/api/overview?view=excel-js-preview)для надстроек Office.

## <a name="common-classes"></a>Общие классы

Приведенный ниже список разделяет основные принципы объектной модели сценариев Office. Здесь показаны общие классы и их связь друг с другом.

- [Рабочая книга](/javascript/api/office-scripts/excelscript/excelscript.workbook) содержит одну или несколько [рабочих листов](/javascript/api/office-scripts/excelscript/excelscript.worksheet).
- [Рабочий лист](/javascript/api/office-scripts/excelscript/excelscript.worksheet) предоставляет доступ к ячейкам через объекты [Range](/javascript/api/office-scripts/excelscript/excelscript.range).
- [Range](/javascript/api/office-scripts/excelscript/excelscript.range) представляет группу смежных клеток.
- [Диапазоны](/javascript/api/office-scripts/excelscript/excelscript.range) используются для создания и размещения [таблиц](/javascript/api/office-scripts/excelscript/excelscript.table), [диаграмм](/javascript/api/office-scripts/excelscript/excelscript.chart), [фигур](/javascript/api/office-scripts/excelscript/excelscript.shape) и других объектов визуализации данных или организации.
- [Лист](/javascript/api/office-scripts/excelscript/excelscript.worksheet) содержит массивы, заполненные объектами, присутствующими на отдельном листе.
- [Книга](/javascript/api/office-scripts/excelscript/excelscript.workbook) содержит массивы некоторых из этих объектов данных для всей книги.

Дополнительные сведения об объектной модели сценариев Office можно найти [в разделе Основные сведения о сценариях для сценариев Office в Excel в Интернете](/office/dev/scripts/develop/scripting-fundamentals)

## <a name="see-also"></a>См. также

- [Сведения о сценариях Office](/office/dev/scripts/overview/excel)
- [Запись, редактирование и создание сценариев Office в Excel в Интернете](/office/dev/scripts/tutorials/excel-tutorial)
- [Основы сценариев для сценариев Office в Excel в Интернете](/office/dev/scripts/develop/scripting-fundamentals)
