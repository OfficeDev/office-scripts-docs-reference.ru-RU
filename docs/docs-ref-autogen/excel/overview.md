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
# <a name="office-scripts-api-reference"></a><span data-ttu-id="d166f-103">Справочник API для сценариев Office</span><span class="sxs-lookup"><span data-stu-id="d166f-103">Office Scripts API reference</span></span>

<span data-ttu-id="d166f-104">API сценариев Office позволяет автоматизировать выполнение распространенных задач в Excel в Интернете.</span><span class="sxs-lookup"><span data-stu-id="d166f-104">The Office Scripts API lets you automate common tasks in Excel on the web.</span></span> <span data-ttu-id="d166f-105">Используйте эту справочную документацию, чтобы узнать больше о классах, методах и других типах, доступных для ваших сценариев.</span><span class="sxs-lookup"><span data-stu-id="d166f-105">Use this reference documentation to learn more about the classes, methods, and other types available for your scripts.</span></span> <span data-ttu-id="d166f-106">Все объекты, доступные с помощью скриптов Office, можно найти в оглавлении слева от страницы.</span><span class="sxs-lookup"><span data-stu-id="d166f-106">All the objects accessible through Office Scripts can be found in the table of contents on the left of the page.</span></span>

> [!NOTE]
> <span data-ttu-id="d166f-107">Если вы ищете API JavaScript для разработки надстроек Office, посетите [Справочник по API JavaScript](/javascript/api/overview?view=excel-js-preview)для надстроек Office.</span><span class="sxs-lookup"><span data-stu-id="d166f-107">If you're looking for the JavaScript APIs for developing Office Add-ins, visit the [Office Add-ins JavaScript API reference](/javascript/api/overview?view=excel-js-preview).</span></span>

## <a name="common-classes"></a><span data-ttu-id="d166f-108">Общие классы</span><span class="sxs-lookup"><span data-stu-id="d166f-108">Common classes</span></span>

<span data-ttu-id="d166f-109">Приведенный ниже список разделяет основные принципы объектной модели сценариев Office.</span><span class="sxs-lookup"><span data-stu-id="d166f-109">The following list breaks down the basics of the Office Scripts object model.</span></span> <span data-ttu-id="d166f-110">Здесь показаны общие классы и их связь друг с другом.</span><span class="sxs-lookup"><span data-stu-id="d166f-110">This shows the common classes and how they relate to one another.</span></span>

- <span data-ttu-id="d166f-111">[Рабочая книга](/javascript/api/office-scripts/excel/excelscript.workbook) содержит одну или несколько [рабочих листов](/javascript/api/office-scripts/excel/excelscript.worksheet).</span><span class="sxs-lookup"><span data-stu-id="d166f-111">A [Workbook](/javascript/api/office-scripts/excel/excelscript.workbook) contains one or more [Worksheets](/javascript/api/office-scripts/excel/excelscript.worksheet).</span></span>
- <span data-ttu-id="d166f-112">[Рабочий лист](/javascript/api/office-scripts/excel/excelscript.worksheet) предоставляет доступ к ячейкам через объекты [Range](/javascript/api/office-scripts/excel/excelscript.range).</span><span class="sxs-lookup"><span data-stu-id="d166f-112">A [Worksheet](/javascript/api/office-scripts/excel/excelscript.worksheet) gives access to cells through [Range](/javascript/api/office-scripts/excel/excelscript.range) objects.</span></span>
- <span data-ttu-id="d166f-113">[Range](/javascript/api/office-scripts/excel/excelscript.range) представляет группу смежных клеток.</span><span class="sxs-lookup"><span data-stu-id="d166f-113">A [Range](/javascript/api/office-scripts/excel/excelscript.range) represents a group of contiguous cells.</span></span>
- <span data-ttu-id="d166f-114">[Диапазоны](/javascript/api/office-scripts/excel/excelscript.range) используются для создания и размещения [таблиц](/javascript/api/office-scripts/excel/excelscript.table), [диаграмм](/javascript/api/office-scripts/excel/excelscript.chart), [фигур](/javascript/api/office-scripts/excel/excelscript.shape) и других объектов визуализации данных или организации.</span><span class="sxs-lookup"><span data-stu-id="d166f-114">[Ranges](/javascript/api/office-scripts/excel/excelscript.range) are used to create and place [Tables](/javascript/api/office-scripts/excel/excelscript.table), [Charts](/javascript/api/office-scripts/excel/excelscript.chart), [Shapes](/javascript/api/office-scripts/excel/excelscript.shape), and other data visualization or organization objects.</span></span>
- <span data-ttu-id="d166f-115">[Лист](/javascript/api/office-scripts/excel/excelscript.worksheet) содержит массивы, заполненные объектами, присутствующими на отдельном листе.</span><span class="sxs-lookup"><span data-stu-id="d166f-115">A [Worksheet](/javascript/api/office-scripts/excel/excelscript.worksheet) contains arrays filled with those objects that are present in the individual sheet.</span></span>
- <span data-ttu-id="d166f-116">[Книга](/javascript/api/office-scripts/excel/excelscript.workbook) содержит массивы некоторых из этих объектов данных для всей книги.</span><span class="sxs-lookup"><span data-stu-id="d166f-116">A [Workbook](/javascript/api/office-scripts/excel/excelscript.workbook) contains arrays of some of those data objects for the entire Workbook.</span></span>

<span data-ttu-id="d166f-117">Дополнительные сведения об объектной модели сценариев Office можно найти [в разделе Основные сведения о сценариях для сценариев Office в Excel в Интернете](/office/dev/scripts/develop/scripting-fundamentals)</span><span class="sxs-lookup"><span data-stu-id="d166f-117">For more information about the Office Scripts object model, visit [Scripting fundamentals for Office Scripts in Excel on the web](/office/dev/scripts/develop/scripting-fundamentals)</span></span>

## <a name="see-also"></a><span data-ttu-id="d166f-118">См. также</span><span class="sxs-lookup"><span data-stu-id="d166f-118">See also</span></span>

- [<span data-ttu-id="d166f-119">Сведения о сценариях Office</span><span class="sxs-lookup"><span data-stu-id="d166f-119">About Office Scripts</span></span>](/office/dev/scripts/overview/excel)
- [<span data-ttu-id="d166f-120">Запись, редактирование и создание сценариев Office в Excel в Интернете</span><span class="sxs-lookup"><span data-stu-id="d166f-120">Record, edit, and create Office Scripts in Excel on the web</span></span>](/office/dev/scripts/tutorials/excel-tutorial)
- [<span data-ttu-id="d166f-121">Основы сценариев для сценариев Office в Excel в Интернете</span><span class="sxs-lookup"><span data-stu-id="d166f-121">Scripting fundamentals for Office Scripts in Excel on the web</span></span>](/office/dev/scripts/develop/scripting-fundamentals)
