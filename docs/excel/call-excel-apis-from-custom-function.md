---
title: Chamar as APIs do Microsoft Excel a partir de uma função personalizada
description: Saiba quais APIs do Microsoft Excel você pode chamar a partir de sua função personalizada.
ms.date: 05/11/2020
localization_priority: Normal
ms.openlocfilehash: a25d3f151f648560ee24a3da3f689cb9767bd52a
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 06/08/2020
ms.locfileid: "44609801"
---
# <a name="call-microsoft-excel-apis-from-a-custom-function"></a><span data-ttu-id="1f5ec-103">Chamar as APIs do Microsoft Excel a partir de uma função personalizada</span><span class="sxs-lookup"><span data-stu-id="1f5ec-103">Call Microsoft Excel APIs from a custom function</span></span>

<span data-ttu-id="1f5ec-104">Chamar as APIs do Excel do Office. js de suas funções personalizadas para obter dados de intervalo e obter mais contexto para seus cálculos.</span><span class="sxs-lookup"><span data-stu-id="1f5ec-104">Call Office.js Excel APIs from your custom functions to get range data and obtain more context for your calculations.</span></span>

<span data-ttu-id="1f5ec-105">Chamar APIs do Office. js por meio de uma função personalizada pode ser útil quando:</span><span class="sxs-lookup"><span data-stu-id="1f5ec-105">Calling Office.js APIs through a custom function can be helpful when:</span></span>

- <span data-ttu-id="1f5ec-106">Uma função personalizada precisa obter informações do Excel antes do cálculo.</span><span class="sxs-lookup"><span data-stu-id="1f5ec-106">A custom function needs to get information from Excel before calculation.</span></span> <span data-ttu-id="1f5ec-107">Essas informações podem incluir propriedades de documento, formatos de intervalo, partes XML personalizadas, um nome de pasta de trabalho ou outras informações específicas do Excel.</span><span class="sxs-lookup"><span data-stu-id="1f5ec-107">This information might include document properties, range formats, custom XML parts, a workbook name, or other Excel-specific information.</span></span>
- <span data-ttu-id="1f5ec-108">Uma função personalizada definirá o formato de número da célula para os valores de retorno após o cálculo.</span><span class="sxs-lookup"><span data-stu-id="1f5ec-108">A custom function will set the cell's number format for the return values after calculation.</span></span>

## <a name="code-sample"></a><span data-ttu-id="1f5ec-109">Exemplo de código</span><span class="sxs-lookup"><span data-stu-id="1f5ec-109">Code sample</span></span>

<span data-ttu-id="1f5ec-110">Para chamar as APIs do Office. js, você precisa primeiro de um contexto.</span><span class="sxs-lookup"><span data-stu-id="1f5ec-110">To call into the Office.js APIs you first need a context.</span></span> <span data-ttu-id="1f5ec-111">Use o `Excel.RequestContext` objeto para obter um contexto.</span><span class="sxs-lookup"><span data-stu-id="1f5ec-111">Use the `Excel.RequestContext` object to get a context.</span></span> <span data-ttu-id="1f5ec-112">Em seguida, use o contexto para chamar as APIs de que você precisa na pasta de trabalho.</span><span class="sxs-lookup"><span data-stu-id="1f5ec-112">Then use the context to call the APIs you need in the workbook.</span></span>

<span data-ttu-id="1f5ec-113">O exemplo de código a seguir mostra como obter um intervalo de valores da pasta de trabalho.</span><span class="sxs-lookup"><span data-stu-id="1f5ec-113">The following code sample shows how to get a range of values from the workbook.</span></span>

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

## <a name="limitations-of-calling-officejs-through-a-custom-function"></a><span data-ttu-id="1f5ec-114">Limitações da chamada do Office. js por meio de uma função personalizada</span><span class="sxs-lookup"><span data-stu-id="1f5ec-114">Limitations of calling Office.js through a custom function</span></span>

<span data-ttu-id="1f5ec-115">Não chame APIs do Office. js de uma função personalizada que altere o ambiente do Excel.</span><span class="sxs-lookup"><span data-stu-id="1f5ec-115">Don't call Office.js APIs from a custom function that change the environment of Excel.</span></span> <span data-ttu-id="1f5ec-116">Isso significa que suas funções personalizadas não devem fazer o seguinte:</span><span class="sxs-lookup"><span data-stu-id="1f5ec-116">This means your custom functions should not do any of the following:</span></span>

- <span data-ttu-id="1f5ec-117">Inserir, excluir ou Formatar células na planilha.</span><span class="sxs-lookup"><span data-stu-id="1f5ec-117">Insert, delete, or format cells on the spreadsheet.</span></span>
- <span data-ttu-id="1f5ec-118">Altera o valor de outra célula.</span><span class="sxs-lookup"><span data-stu-id="1f5ec-118">Change another cell's value.</span></span>
- <span data-ttu-id="1f5ec-119">Mover, renomear, excluir ou adicionar planilhas a uma pasta de trabalho.</span><span class="sxs-lookup"><span data-stu-id="1f5ec-119">Move, rename, delete, or add sheets to a workbook.</span></span>
- <span data-ttu-id="1f5ec-120">Alterar qualquer uma das opções de ambiente, como modo de cálculo ou modos de exibição de tela.</span><span class="sxs-lookup"><span data-stu-id="1f5ec-120">Change any of the environment options, such as calculation mode or screen views.</span></span>
- <span data-ttu-id="1f5ec-121">Adicionar nomes a uma pasta de trabalho.</span><span class="sxs-lookup"><span data-stu-id="1f5ec-121">Add names to a workbook.</span></span>
- <span data-ttu-id="1f5ec-122">Definir propriedades ou executar a maioria dos métodos.</span><span class="sxs-lookup"><span data-stu-id="1f5ec-122">Set properties or execute most methods.</span></span>

<span data-ttu-id="1f5ec-123">Alterar o Excel pode resultar em desempenho ruim, tempo limite e loops infinitos.</span><span class="sxs-lookup"><span data-stu-id="1f5ec-123">Changing Excel can result in poor performance, time outs, and infinite loops.</span></span> <span data-ttu-id="1f5ec-124">Cálculos de função personalizada não devem ser executados enquanto um recálculo do Excel está ocorrendo, pois resultará em resultados imprevisíveis.</span><span class="sxs-lookup"><span data-stu-id="1f5ec-124">Custom function calculations shouldn't run while an Excel recalculation is taking place as it will result in unpredictable results.</span></span>

<span data-ttu-id="1f5ec-125">Em vez disso, faça alterações no Excel a partir do contexto de um botão da faixa de opções ou de um painel de tarefas.</span><span class="sxs-lookup"><span data-stu-id="1f5ec-125">Instead, make changes to Excel from the context of a ribbon button, or task pane.</span></span>

## <a name="next-steps"></a><span data-ttu-id="1f5ec-126">Próximas etapas</span><span class="sxs-lookup"><span data-stu-id="1f5ec-126">Next steps</span></span>

- [<span data-ttu-id="1f5ec-127">Conceitos fundamentais de programação com a API JavaScript do Excel</span><span class="sxs-lookup"><span data-stu-id="1f5ec-127">Fundamental programming concepts with the Excel JavaScript API</span></span>](../reference/overview/excel-add-ins-reference-overview.md)

## <a name="see-also"></a><span data-ttu-id="1f5ec-128">Confira também</span><span class="sxs-lookup"><span data-stu-id="1f5ec-128">See also</span></span>

- [<span data-ttu-id="1f5ec-129">Compartilhar dados e eventos entre as funções personalizadas do Excel e o tutorial do painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="1f5ec-129">Share data and events between Excel custom functions and task pane tutorial</span></span>](../tutorials/share-data-and-events-between-custom-functions-and-the-task-pane-tutorial.md)
