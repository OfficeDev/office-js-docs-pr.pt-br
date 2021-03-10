---
title: Chamar APIs JavaScript do Excel de uma função personalizada
description: Saiba quais APIs JavaScript do Excel você pode chamar de sua função personalizada.
ms.date: 03/05/2021
localization_priority: Normal
ms.openlocfilehash: 4be1b1ee8ea4ae8b2f5d1d27195be18f7aa841da
ms.sourcegitcommit: d153f6d4c3e01d63ed24aa1349be16fa8ad51218
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/10/2021
ms.locfileid: "50613903"
---
# <a name="call-excel-javascript-apis-from-a-custom-function"></a><span data-ttu-id="d87b8-103">Chamar APIs JavaScript do Excel de uma função personalizada</span><span class="sxs-lookup"><span data-stu-id="d87b8-103">Call Excel JavaScript APIs from a custom function</span></span>

<span data-ttu-id="d87b8-104">Chame APIs JavaScript do Excel de suas funções personalizadas para obter dados de intervalo e obter mais contexto para seus cálculos.</span><span class="sxs-lookup"><span data-stu-id="d87b8-104">Call Excel JavaScript APIs from your custom functions to get range data and obtain more context for your calculations.</span></span> <span data-ttu-id="d87b8-105">Chamar APIs JavaScript do Excel por meio de uma função personalizada pode ser útil quando:</span><span class="sxs-lookup"><span data-stu-id="d87b8-105">Calling Excel JavaScript APIs through a custom function can be helpful when:</span></span>

- <span data-ttu-id="d87b8-106">Uma função personalizada precisa obter informações do Excel antes do cálculo.</span><span class="sxs-lookup"><span data-stu-id="d87b8-106">A custom function needs to get information from Excel before calculation.</span></span> <span data-ttu-id="d87b8-107">Essas informações podem incluir propriedades de documento, formatos de intervalo, partes XML personalizadas, um nome da planilha ou outras informações específicas do Excel.</span><span class="sxs-lookup"><span data-stu-id="d87b8-107">This information might include document properties, range formats, custom XML parts, a workbook name, or other Excel-specific information.</span></span>
- <span data-ttu-id="d87b8-108">Uma função personalizada definirá o formato de número da célula para os valores de retorno após o cálculo.</span><span class="sxs-lookup"><span data-stu-id="d87b8-108">A custom function will set the cell's number format for the return values after calculation.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="d87b8-109">Para chamar APIs JavaScript do Excel de sua função personalizada, você precisará usar um tempo de execução JavaScript compartilhado.</span><span class="sxs-lookup"><span data-stu-id="d87b8-109">To call Excel JavaScript APIs from your custom function, you'll need to use a shared JavaScript runtime.</span></span> <span data-ttu-id="d87b8-110">Consulte [Configure seu Suplemento do Office para usar em um tempo de execução do JavaScript compartilhado](../develop/configure-your-add-in-to-use-a-shared-runtime.md) para saber mais.</span><span class="sxs-lookup"><span data-stu-id="d87b8-110">See [Configure your Office Add-in to use a shared JavaScript runtime](../develop/configure-your-add-in-to-use-a-shared-runtime.md) to learn more.</span></span>

## <a name="code-sample"></a><span data-ttu-id="d87b8-111">Exemplo de código</span><span class="sxs-lookup"><span data-stu-id="d87b8-111">Code sample</span></span>

<span data-ttu-id="d87b8-112">Para chamar APIs JavaScript do Excel de uma função personalizada, primeiro você precisa de um contexto.</span><span class="sxs-lookup"><span data-stu-id="d87b8-112">To call Excel JavaScript APIs from a custom function, you first need a context.</span></span> <span data-ttu-id="d87b8-113">Use o [objeto Excel.RequestContext](/javascript/api/excel/excel.requestcontext) para obter um contexto.</span><span class="sxs-lookup"><span data-stu-id="d87b8-113">Use the [Excel.RequestContext](/javascript/api/excel/excel.requestcontext) object to get a context.</span></span> <span data-ttu-id="d87b8-114">Em seguida, use o contexto para chamar as APIs de que você precisa na guia de trabalho.</span><span class="sxs-lookup"><span data-stu-id="d87b8-114">Then use the context to call the APIs you need in the workbook.</span></span>

<span data-ttu-id="d87b8-115">O exemplo de código a seguir mostra como usar para obter um `Excel.RequestContext` valor de uma célula na lista de trabalho.</span><span class="sxs-lookup"><span data-stu-id="d87b8-115">The following code sample shows how to use `Excel.RequestContext` to get a value from a cell in the workbook.</span></span> <span data-ttu-id="d87b8-116">Neste exemplo, o parâmetro é passado para o `address` método [Worksheet.getRange](/javascript/api/excel/excel.worksheet#getRange_address_) da API JavaScript do Excel e deve ser inserido como uma cadeia de caracteres.</span><span class="sxs-lookup"><span data-stu-id="d87b8-116">In this sample, the `address` parameter is passed into the Excel JavaScript API [Worksheet.getRange](/javascript/api/excel/excel.worksheet#getRange_address_) method and must be entered as a string.</span></span> <span data-ttu-id="d87b8-117">Por exemplo, a função personalizada inserida na interface do usuário do Excel deve seguir o padrão , onde é o endereço da célula da qual `=CONTOSO.GETRANGEVALUE("A1")` `"A1"` recuperar o valor.</span><span class="sxs-lookup"><span data-stu-id="d87b8-117">For example, the custom function entered into the Excel UI must follow the pattern `=CONTOSO.GETRANGEVALUE("A1")`, where `"A1"` is the address of the cell from which to retrieve the value.</span></span>

```JavaScript
/**
 * @customfunction
 * @param {string} address The address of the cell from which to retrieve the value.
 * @returns The value of the cell at the input address.
 **/
async function getRangeValue(address) {
 // Retrieve the context object. 
 var context = new Excel.RequestContext();
 
 // Use the context object to access the cell at the input address. 
 var range = context.workbook.worksheets.getActiveWorksheet().getRange(address);
 range.load();
 await context.sync();
 
 // Return the value of the cell at the input address.
 return range.values[0][0];
}
```

## <a name="limitations-of-calling-excel-javascript-apis-through-a-custom-function"></a><span data-ttu-id="d87b8-118">Limitações de chamar APIs JavaScript do Excel por meio de uma função personalizada</span><span class="sxs-lookup"><span data-stu-id="d87b8-118">Limitations of calling Excel JavaScript APIs through a custom function</span></span>

<span data-ttu-id="d87b8-119">Não chame APIs JavaScript do Excel de uma função personalizada que altere o ambiente do Excel.</span><span class="sxs-lookup"><span data-stu-id="d87b8-119">Don't call Excel JavaScript APIs from a custom function that change the environment of Excel.</span></span> <span data-ttu-id="d87b8-120">Isso significa que suas funções personalizadas não devem fazer nada do seguinte:</span><span class="sxs-lookup"><span data-stu-id="d87b8-120">This means your custom functions should not do any of the following:</span></span>

- <span data-ttu-id="d87b8-121">Inserir, excluir ou formatar células na planilha.</span><span class="sxs-lookup"><span data-stu-id="d87b8-121">Insert, delete, or format cells on the spreadsheet.</span></span>
- <span data-ttu-id="d87b8-122">Altere o valor de outra célula.</span><span class="sxs-lookup"><span data-stu-id="d87b8-122">Change another cell's value.</span></span>
- <span data-ttu-id="d87b8-123">Mover, renomear, excluir ou adicionar planilhas a uma planilha.</span><span class="sxs-lookup"><span data-stu-id="d87b8-123">Move, rename, delete, or add sheets to a workbook.</span></span>
- <span data-ttu-id="d87b8-124">Altere qualquer uma das opções de ambiente, como modo de cálculo ou modos de exibição de tela.</span><span class="sxs-lookup"><span data-stu-id="d87b8-124">Change any of the environment options, such as calculation mode or screen views.</span></span>
- <span data-ttu-id="d87b8-125">Adicione nomes a uma lista de trabalho.</span><span class="sxs-lookup"><span data-stu-id="d87b8-125">Add names to a workbook.</span></span>
- <span data-ttu-id="d87b8-126">Definir propriedades ou executar a maioria dos métodos.</span><span class="sxs-lookup"><span data-stu-id="d87b8-126">Set properties or execute most methods.</span></span>

<span data-ttu-id="d87b8-127">Alterar o Excel pode resultar em desempenho ruim, tempo de insufinições e loops infinitos.</span><span class="sxs-lookup"><span data-stu-id="d87b8-127">Changing Excel can result in poor performance, time outs, and infinite loops.</span></span> <span data-ttu-id="d87b8-128">Os cálculos de função personalizada não devem ser executados enquanto um recálculo do Excel está ocorrendo, pois resultará em resultados imprevisíveis.</span><span class="sxs-lookup"><span data-stu-id="d87b8-128">Custom function calculations shouldn't run while an Excel recalculation is taking place as it will result in unpredictable results.</span></span>

<span data-ttu-id="d87b8-129">Em vez disso, faça alterações no Excel a partir do contexto de um botão de faixa de opções ou do painel de tarefas.</span><span class="sxs-lookup"><span data-stu-id="d87b8-129">Instead, make changes to Excel from the context of a ribbon button, or task pane.</span></span>

## <a name="next-steps"></a><span data-ttu-id="d87b8-130">Próximas etapas</span><span class="sxs-lookup"><span data-stu-id="d87b8-130">Next steps</span></span>

- [<span data-ttu-id="d87b8-131">Conceitos fundamentais de programação com a API JavaScript do Excel</span><span class="sxs-lookup"><span data-stu-id="d87b8-131">Fundamental programming concepts with the Excel JavaScript API</span></span>](../reference/overview/excel-add-ins-reference-overview.md)

## <a name="see-also"></a><span data-ttu-id="d87b8-132">Confira também</span><span class="sxs-lookup"><span data-stu-id="d87b8-132">See also</span></span>

- [<span data-ttu-id="d87b8-133">Compartilhar dados e eventos entre funções personalizadas do Excel e tutorial do painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="d87b8-133">Share data and events between Excel custom functions and task pane tutorial</span></span>](../tutorials/share-data-and-events-between-custom-functions-and-the-task-pane-tutorial.md)
- [<span data-ttu-id="d87b8-134">Configure seu Suplemento do Office para usar um tempo de execução de JavaScript compartilhado</span><span class="sxs-lookup"><span data-stu-id="d87b8-134">Configure your Office Add-in to use a shared JavaScript runtime</span></span>](../develop/configure-your-add-in-to-use-a-shared-runtime.md)
