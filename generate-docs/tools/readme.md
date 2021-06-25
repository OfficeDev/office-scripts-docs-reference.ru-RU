# <a name="office-scripts-api-documentation-tools"></a>Office Инструменты документации по API скриптов

Эти средства помогают поддерживать Office SCripts и его команду. Выполните инструкции по запуску средств в этой папке.

## <a name="coverage-tester"></a>coverage-tester

Этот инструмент предоставляет обзор охвата документации для каждого API. Каждый API оценивается для качества документации и наличия примера кода. Показатели качества по-прежнему находятся в разработке.

Выход этого средства — `.csv` файл.

### <a name="coverage-tester-instructions"></a>Инструкции по проверке охвата

1. Клонировать или вить репо.
1. В окне команд перейдите к `/office-scripts-docs-reference/generate-docs/tools`
1. Запуск `npm install`
1. Запуск `npm run build`
1. Запуск `node coverage-tester`
1. Откройте "API Coverage Report.csv"
