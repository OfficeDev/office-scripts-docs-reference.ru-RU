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
# <a name="office-scripts-api-reference"></a><span data-ttu-id="bcf37-103">Справочник API для сценариев Office</span><span class="sxs-lookup"><span data-stu-id="bcf37-103">Office Scripts API reference</span></span>

<span data-ttu-id="bcf37-104">API сценариев Office позволяет автоматизировать выполнение распространенных задач в Excel в Интернете.</span><span class="sxs-lookup"><span data-stu-id="bcf37-104">The Office Scripts API lets you automate common tasks in Excel on the web.</span></span> <span data-ttu-id="bcf37-105">Используйте эту справочную документацию, чтобы узнать больше о классах, методах и других типах, доступных для ваших сценариев.</span><span class="sxs-lookup"><span data-stu-id="bcf37-105">Use this reference documentation to learn more about the classes, methods, and other types available for your scripts.</span></span> <span data-ttu-id="bcf37-106">Все объекты, доступные с помощью скриптов Office, можно найти в оглавлении слева от страницы.</span><span class="sxs-lookup"><span data-stu-id="bcf37-106">All the objects accessible through Office Scripts can be found in the table of contents on the left of the page.</span></span>

> [!NOTE]
> <span data-ttu-id="bcf37-107">Если вы ищете API JavaScript для разработки надстроек Office, посетите [Справочник по API JavaScript](/javascript/api/overview?view=excel-js-preview)для надстроек Office.</span><span class="sxs-lookup"><span data-stu-id="bcf37-107">If you're looking for the JavaScript APIs for developing Office Add-ins, visit the [Office Add-ins JavaScript API reference](/javascript/api/overview?view=excel-js-preview).</span></span>

## <a name="common-classes"></a><span data-ttu-id="bcf37-108">Общие классы</span><span class="sxs-lookup"><span data-stu-id="bcf37-108">Common classes</span></span>

<span data-ttu-id="bcf37-109">Приведенный ниже список разделяет основные принципы объектной модели сценариев Office.</span><span class="sxs-lookup"><span data-stu-id="bcf37-109">The following list breaks down the basics of the Office Scripts object model.</span></span> <span data-ttu-id="bcf37-110">Здесь показаны общие классы и их связь друг с другом.</span><span class="sxs-lookup"><span data-stu-id="bcf37-110">This shows the common classes and how they relate to one another.</span></span>

- <span data-ttu-id="bcf37-111">[Рабочая книга](/javascript/api/office-scripts/excelscript/excelscript.workbook) содержит одну или несколько [рабочих листов](/javascript/api/office-scripts/excelscript/excelscript.worksheet).</span><span class="sxs-lookup"><span data-stu-id="bcf37-111">A [Workbook](/javascript/api/office-scripts/excelscript/excelscript.workbook) contains one or more [Worksheets](/javascript/api/office-scripts/excelscript/excelscript.worksheet).</span></span>
- <span data-ttu-id="bcf37-112">[Рабочий лист](/javascript/api/office-scripts/excelscript/excelscript.worksheet) предоставляет доступ к ячейкам через объекты [Range](/javascript/api/office-scripts/excelscript/excelscript.range).</span><span class="sxs-lookup"><span data-stu-id="bcf37-112">A [Worksheet](/javascript/api/office-scripts/excelscript/excelscript.worksheet) gives access to cells through [Range](/javascript/api/office-scripts/excelscript/excelscript.range) objects.</span></span>
- <span data-ttu-id="bcf37-113">[Range](/javascript/api/office-scripts/excelscript/excelscript.range) представляет группу смежных клеток.</span><span class="sxs-lookup"><span data-stu-id="bcf37-113">A [Range](/javascript/api/office-scripts/excelscript/excelscript.range) represents a group of contiguous cells.</span></span>
- <span data-ttu-id="bcf37-114">[Диапазоны](/javascript/api/office-scripts/excelscript/excelscript.range) используются для создания и размещения [таблиц](/javascript/api/office-scripts/excelscript/excelscript.table), [диаграмм](/javascript/api/office-scripts/excelscript/excelscript.chart), [фигур](/javascript/api/office-scripts/excelscript/excelscript.shape) и других объектов визуализации данных или организации.</span><span class="sxs-lookup"><span data-stu-id="bcf37-114">[Ranges](/javascript/api/office-scripts/excelscript/excelscript.range) are used to create and place [Tables](/javascript/api/office-scripts/excelscript/excelscript.table), [Charts](/javascript/api/office-scripts/excelscript/excelscript.chart), [Shapes](/javascript/api/office-scripts/excelscript/excelscript.shape), and other data visualization or organization objects.</span></span>
- <span data-ttu-id="bcf37-115">[Лист](/javascript/api/office-scripts/excelscript/excelscript.worksheet) содержит массивы, заполненные объектами, присутствующими на отдельном листе.</span><span class="sxs-lookup"><span data-stu-id="bcf37-115">A [Worksheet](/javascript/api/office-scripts/excelscript/excelscript.worksheet) contains arrays filled with those objects that are present in the individual sheet.</span></span>
- <span data-ttu-id="bcf37-116">[Книга](/javascript/api/office-scripts/excelscript/excelscript.workbook) содержит массивы некоторых из этих объектов данных для всей книги.</span><span class="sxs-lookup"><span data-stu-id="bcf37-116">A [Workbook](/javascript/api/office-scripts/excelscript/excelscript.workbook) contains arrays of some of those data objects for the entire Workbook.</span></span>

<span data-ttu-id="bcf37-117">Дополнительные сведения об объектной модели сценариев Office можно найти [в разделе Основные сведения о сценариях для сценариев Office в Excel в Интернете](/office/dev/scripts/develop/scripting-fundamentals)</span><span class="sxs-lookup"><span data-stu-id="bcf37-117">For more information about the Office Scripts object model, visit [Scripting fundamentals for Office Scripts in Excel on the web](/office/dev/scripts/develop/scripting-fundamentals)</span></span>

## <a name="see-also"></a><span data-ttu-id="bcf37-118">См. также</span><span class="sxs-lookup"><span data-stu-id="bcf37-118">See also</span></span>

- [<span data-ttu-id="bcf37-119">Сведения о сценариях Office</span><span class="sxs-lookup"><span data-stu-id="bcf37-119">About Office Scripts</span></span>](/office/dev/scripts/overview/excel)
- [<span data-ttu-id="bcf37-120">Запись, редактирование и создание сценариев Office в Excel в Интернете</span><span class="sxs-lookup"><span data-stu-id="bcf37-120">Record, edit, and create Office Scripts in Excel on the web</span></span>](/office/dev/scripts/tutorials/excel-tutorial)
- [<span data-ttu-id="bcf37-121">Основы сценариев для сценариев Office в Excel в Интернете</span><span class="sxs-lookup"><span data-stu-id="bcf37-121">Scripting fundamentals for Office Scripts in Excel on the web</span></span>](/office/dev/scripts/develop/scripting-fundamentals)
