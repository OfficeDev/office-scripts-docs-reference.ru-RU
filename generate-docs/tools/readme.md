# <a name="office-scripts-api-documentation-tools"></a><span data-ttu-id="b24a5-101">Office Инструменты документации по API скриптов</span><span class="sxs-lookup"><span data-stu-id="b24a5-101">Office Scripts API Documentation Tools</span></span>

<span data-ttu-id="b24a5-102">Эти средства помогают поддерживать Office SCripts и его команду.</span><span class="sxs-lookup"><span data-stu-id="b24a5-102">These tools help support the Office SCripts documentation and the team behind it.</span></span> <span data-ttu-id="b24a5-103">Выполните инструкции по запуску средств в этой папке.</span><span class="sxs-lookup"><span data-stu-id="b24a5-103">Follow these instructions to run the tools in this folder.</span></span>

## <a name="coverage-tester"></a><span data-ttu-id="b24a5-104">coverage-tester</span><span class="sxs-lookup"><span data-stu-id="b24a5-104">coverage-tester</span></span>

<span data-ttu-id="b24a5-105">Этот инструмент предоставляет обзор охвата документации для каждого API.</span><span class="sxs-lookup"><span data-stu-id="b24a5-105">This tool gives an overview of the documentation coverage for each API.</span></span> <span data-ttu-id="b24a5-106">Каждый API оценивается для качества документации и наличия примера кода.</span><span class="sxs-lookup"><span data-stu-id="b24a5-106">Each API is assessed for documentation quality and the presence of sample code.</span></span> <span data-ttu-id="b24a5-107">Показатели качества по-прежнему находятся в разработке.</span><span class="sxs-lookup"><span data-stu-id="b24a5-107">The quality metrics are still in development.</span></span>

<span data-ttu-id="b24a5-108">Выход этого средства — `.csv` файл.</span><span class="sxs-lookup"><span data-stu-id="b24a5-108">The output of this tool is a `.csv` file.</span></span>

### <a name="coverage-tester-instructions"></a><span data-ttu-id="b24a5-109">Инструкции по проверке охвата</span><span class="sxs-lookup"><span data-stu-id="b24a5-109">coverage-tester Instructions</span></span>

1. <span data-ttu-id="b24a5-110">Клонировать или вить репо.</span><span class="sxs-lookup"><span data-stu-id="b24a5-110">Clone or fork the repo.</span></span>
1. <span data-ttu-id="b24a5-111">В окне команд перейдите к `/office-scripts-docs-reference/generate-docs/tools`</span><span class="sxs-lookup"><span data-stu-id="b24a5-111">In a command window, go to `/office-scripts-docs-reference/generate-docs/tools`</span></span>
1. <span data-ttu-id="b24a5-112">Запуск `npm install`</span><span class="sxs-lookup"><span data-stu-id="b24a5-112">Run `npm install`</span></span>
1. <span data-ttu-id="b24a5-113">Запуск `npm run build`</span><span class="sxs-lookup"><span data-stu-id="b24a5-113">Run `npm run build`</span></span>
1. <span data-ttu-id="b24a5-114">Запуск `node coverage-tester`</span><span class="sxs-lookup"><span data-stu-id="b24a5-114">Run `node coverage-tester`</span></span>
1. <span data-ttu-id="b24a5-115">Откройте "API Coverage Report.csv"</span><span class="sxs-lookup"><span data-stu-id="b24a5-115">Open “API Coverage Report.csv”</span></span>
