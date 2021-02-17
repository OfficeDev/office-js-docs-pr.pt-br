---
title: Solução de problemas de complementos do Excel
description: Saiba como solucionar erros de desenvolvimento em Complementos do Excel.
ms.date: 02/12/2021
localization_priority: Normal
ms.openlocfilehash: 0efc8b4d25d9d748975146e187104972e4ad58a9
ms.sourcegitcommit: 1cdf5728102424a46998e1527508b4e7f9f74a4c
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 02/17/2021
ms.locfileid: "50270725"
---
# <a name="troubleshooting-excel-add-ins"></a><span data-ttu-id="53a06-103">Solução de problemas de complementos do Excel</span><span class="sxs-lookup"><span data-stu-id="53a06-103">Troubleshooting Excel Add-ins</span></span>

<span data-ttu-id="53a06-104">Este artigo discute a solução de problemas que são exclusivos do Excel.</span><span class="sxs-lookup"><span data-stu-id="53a06-104">This article discusses troubleshooting issues that are unique to Excel.</span></span> <span data-ttu-id="53a06-105">Use a ferramenta de comentários na parte inferior da página para sugerir outros problemas que podem ser adicionados ao artigo.</span><span class="sxs-lookup"><span data-stu-id="53a06-105">Please use the feedback tool at the bottom of the page to suggest other issues that can be added to the article.</span></span>

## <a name="api-limitations-when-the-active-workbook-switches"></a><span data-ttu-id="53a06-106">Limitações de API quando a agenda ativa é alternada</span><span class="sxs-lookup"><span data-stu-id="53a06-106">API limitations when the active workbook switches</span></span>

<span data-ttu-id="53a06-107">Os complementos do Excel se destinam a operar em uma única planilha de cada vez.</span><span class="sxs-lookup"><span data-stu-id="53a06-107">Add-ins for Excel are intended to operate on a single workbook at a time.</span></span> <span data-ttu-id="53a06-108">Erros podem surgir quando uma área de trabalho separada da que está executando o complemento ganha foco.</span><span class="sxs-lookup"><span data-stu-id="53a06-108">Errors can arise when a workbook that is separate from the one running the add-in gains focus.</span></span> <span data-ttu-id="53a06-109">Isso só acontece quando métodos específicos estão sendo chamados quando o foco muda.</span><span class="sxs-lookup"><span data-stu-id="53a06-109">This only happens when particular methods are in the process of being called when the focus changes.</span></span>

<span data-ttu-id="53a06-110">As seguintes APIs são afetadas por essa opção de livro de trabalho:</span><span class="sxs-lookup"><span data-stu-id="53a06-110">The following APIs are affected by this workbook switch:</span></span>

|<span data-ttu-id="53a06-111">API JavaScript do Excel</span><span class="sxs-lookup"><span data-stu-id="53a06-111">Excel JavaScript API</span></span> | <span data-ttu-id="53a06-112">Erro lançado</span><span class="sxs-lookup"><span data-stu-id="53a06-112">Error thrown</span></span> |
|--|--|
| `Chart.activate` | <span data-ttu-id="53a06-113">GeneralException</span><span class="sxs-lookup"><span data-stu-id="53a06-113">GeneralException</span></span> |
| `Range.select` | <span data-ttu-id="53a06-114">GeneralException</span><span class="sxs-lookup"><span data-stu-id="53a06-114">GeneralException</span></span> |
| `Table.clearFilters` | <span data-ttu-id="53a06-115">GeneralException</span><span class="sxs-lookup"><span data-stu-id="53a06-115">GeneralException</span></span> |
| `Workbook.getActiveCell`  | <span data-ttu-id="53a06-116">InvalidSelection</span><span class="sxs-lookup"><span data-stu-id="53a06-116">InvalidSelection</span></span>|
| `Workbook.getSelectedRange` | <span data-ttu-id="53a06-117">InvalidSelection</span><span class="sxs-lookup"><span data-stu-id="53a06-117">InvalidSelection</span></span>|
| `Workbook.getSelectedRanges`  | <span data-ttu-id="53a06-118">InvalidSelection</span><span class="sxs-lookup"><span data-stu-id="53a06-118">InvalidSelection</span></span>|
| `Worksheet.activate` | <span data-ttu-id="53a06-119">GeneralException</span><span class="sxs-lookup"><span data-stu-id="53a06-119">GeneralException</span></span> |
| `Worksheet.delete`  | <span data-ttu-id="53a06-120">InvalidSelection</span><span class="sxs-lookup"><span data-stu-id="53a06-120">InvalidSelection</span></span>|
| `Worksheet.gridlines` | <span data-ttu-id="53a06-121">GeneralException</span><span class="sxs-lookup"><span data-stu-id="53a06-121">GeneralException</span></span> |
| `Worksheet.showHeadings` | <span data-ttu-id="53a06-122">GeneralException</span><span class="sxs-lookup"><span data-stu-id="53a06-122">GeneralException</span></span> |
| `WorksheetCollection.add` | <span data-ttu-id="53a06-123">GeneralException</span><span class="sxs-lookup"><span data-stu-id="53a06-123">GeneralException</span></span> |
| `WorksheetFreezePanes.freezeAt` | <span data-ttu-id="53a06-124">GeneralException</span><span class="sxs-lookup"><span data-stu-id="53a06-124">GeneralException</span></span> |
| `WorksheetFreezePanes.freezeColumns` | <span data-ttu-id="53a06-125">GeneralException</span><span class="sxs-lookup"><span data-stu-id="53a06-125">GeneralException</span></span> |
| `WorksheetFreezePanes.freezeRows` | <span data-ttu-id="53a06-126">GeneralException</span><span class="sxs-lookup"><span data-stu-id="53a06-126">GeneralException</span></span> |
| `WorksheetFreezePanes.getLocationOrNullObject`| <span data-ttu-id="53a06-127">GeneralException</span><span class="sxs-lookup"><span data-stu-id="53a06-127">GeneralException</span></span> |
| `WorksheetFreezePanes.unfreeze` | <span data-ttu-id="53a06-128">GeneralException</span><span class="sxs-lookup"><span data-stu-id="53a06-128">GeneralException</span></span> |

> [!NOTE]
> <span data-ttu-id="53a06-129">Isso só se aplica a várias planilhas do Excel abertas no Windows ou Mac.</span><span class="sxs-lookup"><span data-stu-id="53a06-129">This only applies to multiple Excel workbooks open on Windows or Mac.</span></span>

## <a name="coauthoring"></a><span data-ttu-id="53a06-130">Coautoria</span><span class="sxs-lookup"><span data-stu-id="53a06-130">Coauthoring</span></span>

<span data-ttu-id="53a06-131">Confira [Coautor nos complementos do Excel](co-authoring-in-excel-add-ins.md) para padrões a usar com eventos em um ambiente de coautor.</span><span class="sxs-lookup"><span data-stu-id="53a06-131">See [Coauthoring in Excel add-ins](co-authoring-in-excel-add-ins.md) for patterns to use with events in a coauthoring environment.</span></span> <span data-ttu-id="53a06-132">O artigo também discute possíveis conflitos de mesclagem ao usar determinadas APIs, como [`TableRowCollection.add`](/javascript/api/excel/excel.tablerowcollection#add-index--values-) .</span><span class="sxs-lookup"><span data-stu-id="53a06-132">The article also discusses potential merge conflicts when using certain APIs, such as [`TableRowCollection.add`](/javascript/api/excel/excel.tablerowcollection#add-index--values-).</span></span>

## <a name="known-issues"></a><span data-ttu-id="53a06-133">Problemas Conhecidos</span><span class="sxs-lookup"><span data-stu-id="53a06-133">Known Issues</span></span>

### <a name="binding-events-return-temporary-binding-obects"></a><span data-ttu-id="53a06-134">Eventos de associação `Binding` retornam obects temporários</span><span class="sxs-lookup"><span data-stu-id="53a06-134">Binding events return temporary `Binding` obects</span></span>

<span data-ttu-id="53a06-135">[BindingDataChangedEventArgs.binding](/javascript/api/excel/excel.bindingdatachangedeventargs#binding) e [BindingSelectionChangedEventArgs.binding](/javascript/api/excel/excel.bindingselectionchangedeventargs#binding) retornam um objeto temporário que contém a ID do objeto que gerou o `Binding` `Binding` evento.</span><span class="sxs-lookup"><span data-stu-id="53a06-135">Both [BindingDataChangedEventArgs.binding](/javascript/api/excel/excel.bindingdatachangedeventargs#binding) and [BindingSelectionChangedEventArgs.binding](/javascript/api/excel/excel.bindingselectionchangedeventargs#binding) return a temporary `Binding` object that contains the ID of the `Binding` object that raised the event.</span></span> <span data-ttu-id="53a06-136">Use essa ID com `BindingCollection.getItem(id)` para recuperar o objeto que gerou o `Binding` evento.</span><span class="sxs-lookup"><span data-stu-id="53a06-136">Use this ID with `BindingCollection.getItem(id)` to retrieve the `Binding` object that raised the event.</span></span>

<span data-ttu-id="53a06-137">O exemplo de código a seguir mostra como usar essa ID de associação temporária para recuperar o objeto `Binding` relacionado.</span><span class="sxs-lookup"><span data-stu-id="53a06-137">The following code sample shows how to use this temporary binding ID to retrieve the related `Binding` object.</span></span> <span data-ttu-id="53a06-138">No exemplo, um ouvinte de eventos é atribuído a uma associação.</span><span class="sxs-lookup"><span data-stu-id="53a06-138">In the sample, an event listener is assigned to a binding.</span></span> <span data-ttu-id="53a06-139">O ouvinte chama `getBindingId` o método quando o evento é `onDataChanged` disparado.</span><span class="sxs-lookup"><span data-stu-id="53a06-139">The listener calls the `getBindingId` method when the `onDataChanged` event is triggered.</span></span> <span data-ttu-id="53a06-140">O `getBindingId` método usa a ID do objeto temporário para recuperar o objeto que gerou o `Binding` `Binding` evento.</span><span class="sxs-lookup"><span data-stu-id="53a06-140">The `getBindingId` method uses the ID of the temporary `Binding` object to retrieve the `Binding` object that raised the event.</span></span>

```js
Excel.run(function (context) {
    // Retrieve your binding.
    var binding = context.workbook.bindings.getItemAt(0);

    return context.sync().then(function () {
        // Register an event listener to detect changes to your binding
        // and then trigger the `getBindingId` method when the data changes. 
        binding.onDataChanged.add(getBindingId);

        return context.sync();
    });
});

function getBindingId(eventArgs) {
    return Excel.run(function (context) {
        // Get the temporary binding object and load its ID. 
        var tempBindingObject = eventArgs.binding;
        tempBindingObject.load("id");

        // Use the temporary binding object's ID to retrieve the original binding object. 
        var originalBindingObject = context.workbook.bindings.getItem(tempBindingObject.id);

        // You now have the binding object that raised the event: `originalBindingObject`. 
    });
}
```

### <a name="cell-format-usestandardheight-and-usestandardwidth-issues"></a><span data-ttu-id="53a06-141">Formato e `useStandardHeight` problemas da `useStandardWidth` célula</span><span class="sxs-lookup"><span data-stu-id="53a06-141">Cell format `useStandardHeight` and `useStandardWidth` issues</span></span>

<span data-ttu-id="53a06-142">A [propriedade useStandardHeight](/javascript/api/excel/excel.cellpropertiesformat#useStandardHeight) `CellPropertiesFormat` não funciona corretamente no Excel na Web.</span><span class="sxs-lookup"><span data-stu-id="53a06-142">The [useStandardHeight](/javascript/api/excel/excel.cellpropertiesformat#useStandardHeight) property of `CellPropertiesFormat` doesn't work properly in Excel on the web.</span></span> <span data-ttu-id="53a06-143">Devido a um problema na interface do usuário do Excel na Web, definir a propriedade para calcular a altura de forma `useStandardHeight` `true` imprecisa nessa plataforma.</span><span class="sxs-lookup"><span data-stu-id="53a06-143">Due to an issue in the Excel on the web UI, setting the `useStandardHeight` property to `true` calculates height imprecisely on this platform.</span></span> <span data-ttu-id="53a06-144">Por exemplo, uma altura padrão **de 14 é** modificada para **14,25** no Excel na Web.</span><span class="sxs-lookup"><span data-stu-id="53a06-144">For example, a standard height of **14** is modified to **14.25** in Excel on the web.</span></span>

<span data-ttu-id="53a06-145">Em todas as plataformas, as propriedades [useStandardHeight](/javascript/api/excel/excel.cellpropertiesformat#useStandardHeight) e [useStandardWidth](/javascript/api/excel/excel.cellpropertiesformat#useStandardWidth) devem `CellPropertiesFormat` ser definidas apenas como `true` .</span><span class="sxs-lookup"><span data-stu-id="53a06-145">On all platforms, the [useStandardHeight](/javascript/api/excel/excel.cellpropertiesformat#useStandardHeight) and [useStandardWidth](/javascript/api/excel/excel.cellpropertiesformat#useStandardWidth) properties of `CellPropertiesFormat` are only intended to be set to `true`.</span></span> <span data-ttu-id="53a06-146">A definição dessas propriedades `false` não tem efeito.</span><span class="sxs-lookup"><span data-stu-id="53a06-146">Setting these properties to `false` has no effect.</span></span> 

### <a name="range-getimage-method-unsupported-on-excel-for-mac"></a><span data-ttu-id="53a06-147">Método `getImage` Range sem suporte no Excel para Mac</span><span class="sxs-lookup"><span data-stu-id="53a06-147">Range `getImage` method unsupported on Excel for Mac</span></span>

<span data-ttu-id="53a06-148">O método [GetImage](/javascript/api/excel/excel.range#getImage__) de intervalo não tem suporte no Excel para Mac no momento.</span><span class="sxs-lookup"><span data-stu-id="53a06-148">The Range [getImage](/javascript/api/excel/excel.range#getImage__) method isn't currently supported in Excel for Mac.</span></span> <span data-ttu-id="53a06-149">Consulte [o problema OfficeDev/office-js #235](https://github.com/OfficeDev/office-js/issues/235) para o status atual.</span><span class="sxs-lookup"><span data-stu-id="53a06-149">See [OfficeDev/office-js Issue #235](https://github.com/OfficeDev/office-js/issues/235) for the current status.</span></span>

### <a name="range-return-character-limit"></a><span data-ttu-id="53a06-150">Limite de caracteres de retorno de intervalo</span><span class="sxs-lookup"><span data-stu-id="53a06-150">Range return character limit</span></span>

<span data-ttu-id="53a06-151">Os [métodos Worksheet.getRange(address)](/javascript/api/excel/excel.worksheet#getRange_address_) e [Worksheet.getRanges(address)](/javascript/api/excel/excel.worksheet#getRanges_address_) têm um limite de cadeia de caracteres de endereço de 8192 caracteres.</span><span class="sxs-lookup"><span data-stu-id="53a06-151">The [Worksheet.getRange(address)](/javascript/api/excel/excel.worksheet#getRange_address_) and [Worksheet.getRanges(address)](/javascript/api/excel/excel.worksheet#getRanges_address_) methods have an address string limit of 8192 characters.</span></span> <span data-ttu-id="53a06-152">Quando esse limite é excedido, a cadeia de caracteres do endereço é truncada para 8192 caracteres.</span><span class="sxs-lookup"><span data-stu-id="53a06-152">When this limit is exceeded, the address string is truncated to 8192 characters.</span></span>

## <a name="see-also"></a><span data-ttu-id="53a06-153">Confira também</span><span class="sxs-lookup"><span data-stu-id="53a06-153">See also</span></span>

- [<span data-ttu-id="53a06-154">Solucionar erros de desenvolvimento com os Complementos do Office</span><span class="sxs-lookup"><span data-stu-id="53a06-154">Troubleshoot development errors with Office Add-ins</span></span>](../testing/troubleshoot-development-errors.md)
- [<span data-ttu-id="53a06-155">Solucionar erros de usuários com Suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="53a06-155">Troubleshoot user errors with Office Add-ins</span></span>](../testing/testing-and-troubleshooting.md)
