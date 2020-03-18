---
title: Chamar as APIs do Microsoft Excel a partir de uma função personalizada
description: Saiba quais APIs do Microsoft Excel você pode chamar a partir de sua função personalizada.
ms.date: 02/06/2020
localization_priority: Normal
ms.openlocfilehash: e22ed897e95a74707bd0d8bded3f8dca724731d1
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/17/2020
ms.locfileid: "42719340"
---
# <a name="call-microsoft-excel-apis-from-a-custom-function"></a><span data-ttu-id="ac699-103">Chamar as APIs do Microsoft Excel a partir de uma função personalizada</span><span class="sxs-lookup"><span data-stu-id="ac699-103">Call Microsoft Excel APIs from a custom function</span></span>

[!include[Running custom functions in a shared runtime note](../includes/excel-shared-runtime-preview-note.md)]

<span data-ttu-id="ac699-104">Chamar as APIs do Excel do Office. js de suas funções personalizadas para obter dados de intervalo e obter mais contexto para seus cálculos.</span><span class="sxs-lookup"><span data-stu-id="ac699-104">Call Office.js Excel APIs from your custom functions to get range data and obtain more context for your calculations.</span></span>

<span data-ttu-id="ac699-105">Chamar APIs do Office. js por meio de uma função personalizada pode ser útil quando:</span><span class="sxs-lookup"><span data-stu-id="ac699-105">Calling Office.js APIs through a custom function can be helpful when:</span></span>

- <span data-ttu-id="ac699-106">Uma função personalizada precisa obter informações do Excel antes do cálculo.</span><span class="sxs-lookup"><span data-stu-id="ac699-106">A custom function needs to get information from Excel before calculation.</span></span> <span data-ttu-id="ac699-107">Essas informações podem incluir propriedades de documento, formatos de intervalo, partes XML personalizadas, um nome de pasta de trabalho ou outras informações específicas do Excel.</span><span class="sxs-lookup"><span data-stu-id="ac699-107">This information might include document properties, range formats, custom XML parts, a workbook name, or other Excel-specific information.</span></span>
- <span data-ttu-id="ac699-108">Uma função personalizada definirá o formato de número da célula para os valores de retorno após o cálculo.</span><span class="sxs-lookup"><span data-stu-id="ac699-108">A custom function will set the cell's number format for the return values after calculation.</span></span>

[!include[Excel shared runtime note](../includes/note-requires-shared-runtime.md)]

## <a name="code-sample"></a><span data-ttu-id="ac699-109">Exemplo de código</span><span class="sxs-lookup"><span data-stu-id="ac699-109">Code sample</span></span>

<span data-ttu-id="ac699-110">Para chamar as APIs do Office. js, você precisa primeiro de um contexto.</span><span class="sxs-lookup"><span data-stu-id="ac699-110">To call into the Office.js APIs you first need a context.</span></span> <span data-ttu-id="ac699-111">Use o `Excel.RequestContext` objeto para obter um contexto.</span><span class="sxs-lookup"><span data-stu-id="ac699-111">Use the `Excel.RequestContext` object to get a context.</span></span> <span data-ttu-id="ac699-112">Em seguida, use o contexto para chamar as APIs de que você precisa na pasta de trabalho.</span><span class="sxs-lookup"><span data-stu-id="ac699-112">Then use the context to call the APIs you need in the workbook.</span></span>

<span data-ttu-id="ac699-113">O exemplo de código a seguir mostra como obter um intervalo de valores da pasta de trabalho.</span><span class="sxs-lookup"><span data-stu-id="ac699-113">The following code sample shows how to get a range of values from the workbook.</span></span>

```JavaScript
/**
 * @customfunction
 * @param address range's address
 **/
async function getRangeValue (address) {
 var context = new Excel.RequestContext();
 var range = context.workbook.worksheets.getActiveWorksheet().getRange(address);
 range.load();
 await context.sync();
 return range.values[0][0];
}
```

## <a name="limitations-of-calling-officejs-through-a-custom-function"></a><span data-ttu-id="ac699-114">Limitações da chamada do Office. js por meio de uma função personalizada</span><span class="sxs-lookup"><span data-stu-id="ac699-114">Limitations of calling Office.js through a custom function</span></span>

<span data-ttu-id="ac699-115">Não chame APIs do Office. js de uma função personalizada que altere o ambiente do Excel.</span><span class="sxs-lookup"><span data-stu-id="ac699-115">Don't call Office.js APIs from a custom function that change the environment of Excel.</span></span> <span data-ttu-id="ac699-116">Isso significa que suas funções personalizadas não devem fazer o seguinte:</span><span class="sxs-lookup"><span data-stu-id="ac699-116">This means your custom functions should not do any of the following:</span></span>

- <span data-ttu-id="ac699-117">Inserir, excluir ou Formatar células na planilha.</span><span class="sxs-lookup"><span data-stu-id="ac699-117">Insert, delete, or format cells on the spreadsheet.</span></span>
- <span data-ttu-id="ac699-118">Altera o valor de outra célula.</span><span class="sxs-lookup"><span data-stu-id="ac699-118">Change another cell's value.</span></span>
- <span data-ttu-id="ac699-119">Mover, renomear, excluir ou adicionar planilhas a uma pasta de trabalho.</span><span class="sxs-lookup"><span data-stu-id="ac699-119">Move, rename, delete, or add sheets to a workbook.</span></span>
- <span data-ttu-id="ac699-120">Alterar qualquer uma das opções de ambiente, como modo de cálculo ou modos de exibição de tela.</span><span class="sxs-lookup"><span data-stu-id="ac699-120">Change any of the environment options, such as calculation mode or screen views.</span></span>
- <span data-ttu-id="ac699-121">Adicionar nomes a uma pasta de trabalho.</span><span class="sxs-lookup"><span data-stu-id="ac699-121">Add names to a workbook.</span></span>
- <span data-ttu-id="ac699-122">Definir propriedades ou executar a maioria dos métodos.</span><span class="sxs-lookup"><span data-stu-id="ac699-122">Set properties or execute most methods.</span></span>

<span data-ttu-id="ac699-123">Alterar o Excel pode resultar em desempenho ruim, tempo limite e loops infinitos.</span><span class="sxs-lookup"><span data-stu-id="ac699-123">Changing Excel can result in poor performance, time outs, and infinite loops.</span></span> <span data-ttu-id="ac699-124">Cálculos de função personalizada não devem ser executados enquanto um recálculo do Excel está ocorrendo, pois resultará em resultados imprevisíveis.</span><span class="sxs-lookup"><span data-stu-id="ac699-124">Custom function calculations shouldn't run while an Excel recalculation is taking place as it will result in unpredictable results.</span></span>

<span data-ttu-id="ac699-125">Em vez disso, faça alterações no Excel a partir do contexto de um botão da faixa de opções ou de um painel de tarefas.</span><span class="sxs-lookup"><span data-stu-id="ac699-125">Instead, make changes to Excel from the context of a ribbon button, or task pane.</span></span>

## <a name="next-steps"></a><span data-ttu-id="ac699-126">Próximas etapas</span><span class="sxs-lookup"><span data-stu-id="ac699-126">Next steps</span></span>

- [<span data-ttu-id="ac699-127">Conceitos fundamentais de programação com a API JavaScript do Excel</span><span class="sxs-lookup"><span data-stu-id="ac699-127">Fundamental programming concepts with the Excel JavaScript API</span></span>](../reference/overview/excel-add-ins-reference-overview.md)

## <a name="see-also"></a><span data-ttu-id="ac699-128">Também confira</span><span class="sxs-lookup"><span data-stu-id="ac699-128">See also</span></span>

- [<span data-ttu-id="ac699-129">Compartilhar dados e eventos entre as funções personalizadas do Excel e o tutorial do painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="ac699-129">Share data and events between Excel custom functions and task pane tutorial</span></span>](../tutorials/share-data-and-events-between-custom-functions-and-the-task-pane-tutorial.md)