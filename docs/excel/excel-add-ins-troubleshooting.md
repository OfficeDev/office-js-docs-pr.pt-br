---
title: Solução de Excel de solução de problemas
description: Saiba como solucionar erros de desenvolvimento em Excel de complementos.
ms.date: 02/12/2021
localization_priority: Normal
ms.openlocfilehash: cb622a1805be7bec61168ab37a41709a57075788
ms.sourcegitcommit: 883f71d395b19ccfc6874a0d5942a7016eb49e2c
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 07/09/2021
ms.locfileid: "53349438"
---
# <a name="troubleshooting-excel-add-ins"></a><span data-ttu-id="bb311-103">Solução de Excel de solução de problemas</span><span class="sxs-lookup"><span data-stu-id="bb311-103">Troubleshooting Excel Add-ins</span></span>

<span data-ttu-id="bb311-104">Este artigo discute a solução de problemas que são exclusivos Excel.</span><span class="sxs-lookup"><span data-stu-id="bb311-104">This article discusses troubleshooting issues that are unique to Excel.</span></span> <span data-ttu-id="bb311-105">Use a ferramenta de comentários na parte inferior da página para sugerir outros problemas que podem ser adicionados ao artigo.</span><span class="sxs-lookup"><span data-stu-id="bb311-105">Please use the feedback tool at the bottom of the page to suggest other issues that can be added to the article.</span></span>

## <a name="api-limitations-when-the-active-workbook-switches"></a><span data-ttu-id="bb311-106">Limitações da API quando a agenda de trabalho ativa é alternada</span><span class="sxs-lookup"><span data-stu-id="bb311-106">API limitations when the active workbook switches</span></span>

<span data-ttu-id="bb311-107">Os complementos para Excel são destinados a operar em uma única workbook de cada vez.</span><span class="sxs-lookup"><span data-stu-id="bb311-107">Add-ins for Excel are intended to operate on a single workbook at a time.</span></span> <span data-ttu-id="bb311-108">Erros podem surgir quando uma workbook separada da que está executando o complemento ganha o foco.</span><span class="sxs-lookup"><span data-stu-id="bb311-108">Errors can arise when a workbook that is separate from the one running the add-in gains focus.</span></span> <span data-ttu-id="bb311-109">Isso só acontece quando determinados métodos estão no processo de ser chamado quando o foco muda.</span><span class="sxs-lookup"><span data-stu-id="bb311-109">This only happens when particular methods are in the process of being called when the focus changes.</span></span>

<span data-ttu-id="bb311-110">As APIs a seguir são afetadas por essa opção de lista de trabalho.</span><span class="sxs-lookup"><span data-stu-id="bb311-110">The following APIs are affected by this workbook switch.</span></span>

|<span data-ttu-id="bb311-111">API JavaScript do Excel</span><span class="sxs-lookup"><span data-stu-id="bb311-111">Excel JavaScript API</span></span> | <span data-ttu-id="bb311-112">Erro lançado</span><span class="sxs-lookup"><span data-stu-id="bb311-112">Error thrown</span></span> |
|--|--|
| `Chart.activate` | <span data-ttu-id="bb311-113">GeneralException</span><span class="sxs-lookup"><span data-stu-id="bb311-113">GeneralException</span></span> |
| `Range.select` | <span data-ttu-id="bb311-114">GeneralException</span><span class="sxs-lookup"><span data-stu-id="bb311-114">GeneralException</span></span> |
| `Table.clearFilters` | <span data-ttu-id="bb311-115">GeneralException</span><span class="sxs-lookup"><span data-stu-id="bb311-115">GeneralException</span></span> |
| `Workbook.getActiveCell`  | <span data-ttu-id="bb311-116">InvalidSelection</span><span class="sxs-lookup"><span data-stu-id="bb311-116">InvalidSelection</span></span>|
| `Workbook.getSelectedRange` | <span data-ttu-id="bb311-117">InvalidSelection</span><span class="sxs-lookup"><span data-stu-id="bb311-117">InvalidSelection</span></span>|
| `Workbook.getSelectedRanges`  | <span data-ttu-id="bb311-118">InvalidSelection</span><span class="sxs-lookup"><span data-stu-id="bb311-118">InvalidSelection</span></span>|
| `Worksheet.activate` | <span data-ttu-id="bb311-119">GeneralException</span><span class="sxs-lookup"><span data-stu-id="bb311-119">GeneralException</span></span> |
| `Worksheet.delete`  | <span data-ttu-id="bb311-120">InvalidSelection</span><span class="sxs-lookup"><span data-stu-id="bb311-120">InvalidSelection</span></span>|
| `Worksheet.gridlines` | <span data-ttu-id="bb311-121">GeneralException</span><span class="sxs-lookup"><span data-stu-id="bb311-121">GeneralException</span></span> |
| `Worksheet.showHeadings` | <span data-ttu-id="bb311-122">GeneralException</span><span class="sxs-lookup"><span data-stu-id="bb311-122">GeneralException</span></span> |
| `WorksheetCollection.add` | <span data-ttu-id="bb311-123">GeneralException</span><span class="sxs-lookup"><span data-stu-id="bb311-123">GeneralException</span></span> |
| `WorksheetFreezePanes.freezeAt` | <span data-ttu-id="bb311-124">GeneralException</span><span class="sxs-lookup"><span data-stu-id="bb311-124">GeneralException</span></span> |
| `WorksheetFreezePanes.freezeColumns` | <span data-ttu-id="bb311-125">GeneralException</span><span class="sxs-lookup"><span data-stu-id="bb311-125">GeneralException</span></span> |
| `WorksheetFreezePanes.freezeRows` | <span data-ttu-id="bb311-126">GeneralException</span><span class="sxs-lookup"><span data-stu-id="bb311-126">GeneralException</span></span> |
| `WorksheetFreezePanes.getLocationOrNullObject`| <span data-ttu-id="bb311-127">GeneralException</span><span class="sxs-lookup"><span data-stu-id="bb311-127">GeneralException</span></span> |
| `WorksheetFreezePanes.unfreeze` | <span data-ttu-id="bb311-128">GeneralException</span><span class="sxs-lookup"><span data-stu-id="bb311-128">GeneralException</span></span> |

> [!NOTE]
> <span data-ttu-id="bb311-129">Isso só se aplica a várias Excel de trabalho abertas em Windows ou Mac.</span><span class="sxs-lookup"><span data-stu-id="bb311-129">This only applies to multiple Excel workbooks open on Windows or Mac.</span></span>

## <a name="coauthoring"></a><span data-ttu-id="bb311-130">Coautoria</span><span class="sxs-lookup"><span data-stu-id="bb311-130">Coauthoring</span></span>

<span data-ttu-id="bb311-131">Consulte [Coautor no Excel para](co-authoring-in-excel-add-ins.md) padrões a ser usado com eventos em um ambiente de coautor.</span><span class="sxs-lookup"><span data-stu-id="bb311-131">See [Coauthoring in Excel add-ins](co-authoring-in-excel-add-ins.md) for patterns to use with events in a coauthoring environment.</span></span> <span data-ttu-id="bb311-132">O artigo também aborda possíveis conflitos de mesclagem ao usar determinadas APIs, como [`TableRowCollection.add`](/javascript/api/excel/excel.tablerowcollection#add-index--values-) .</span><span class="sxs-lookup"><span data-stu-id="bb311-132">The article also discusses potential merge conflicts when using certain APIs, such as [`TableRowCollection.add`](/javascript/api/excel/excel.tablerowcollection#add-index--values-).</span></span>

## <a name="known-issues"></a><span data-ttu-id="bb311-133">Problemas Conhecidos</span><span class="sxs-lookup"><span data-stu-id="bb311-133">Known Issues</span></span>

### <a name="binding-events-return-temporary-binding-obects"></a><span data-ttu-id="bb311-134">Eventos de associação `Binding` retornam obects temporários</span><span class="sxs-lookup"><span data-stu-id="bb311-134">Binding events return temporary `Binding` obects</span></span>

<span data-ttu-id="bb311-135">[BindingDataChangedEventArgs.binding](/javascript/api/excel/excel.bindingdatachangedeventargs#binding) e [BindingSelectionChangedEventArgs.binding](/javascript/api/excel/excel.bindingselectionchangedeventargs#binding) retornam um objeto temporário que contém a ID do objeto que gerou o `Binding` `Binding` evento.</span><span class="sxs-lookup"><span data-stu-id="bb311-135">Both [BindingDataChangedEventArgs.binding](/javascript/api/excel/excel.bindingdatachangedeventargs#binding) and [BindingSelectionChangedEventArgs.binding](/javascript/api/excel/excel.bindingselectionchangedeventargs#binding) return a temporary `Binding` object that contains the ID of the `Binding` object that raised the event.</span></span> <span data-ttu-id="bb311-136">Use essa ID com `BindingCollection.getItem(id)` para recuperar o objeto que gerou o `Binding` evento.</span><span class="sxs-lookup"><span data-stu-id="bb311-136">Use this ID with `BindingCollection.getItem(id)` to retrieve the `Binding` object that raised the event.</span></span>

<span data-ttu-id="bb311-137">O exemplo de código a seguir mostra como usar essa ID de associação temporária para recuperar o objeto `Binding` relacionado.</span><span class="sxs-lookup"><span data-stu-id="bb311-137">The following code sample shows how to use this temporary binding ID to retrieve the related `Binding` object.</span></span> <span data-ttu-id="bb311-138">No exemplo, um ouvinte de eventos é atribuído a uma associação.</span><span class="sxs-lookup"><span data-stu-id="bb311-138">In the sample, an event listener is assigned to a binding.</span></span> <span data-ttu-id="bb311-139">O ouvinte chama `getBindingId` o método quando o evento é `onDataChanged` disparado.</span><span class="sxs-lookup"><span data-stu-id="bb311-139">The listener calls the `getBindingId` method when the `onDataChanged` event is triggered.</span></span> <span data-ttu-id="bb311-140">O `getBindingId` método usa a ID do objeto temporário para recuperar o objeto que gerou o `Binding` `Binding` evento.</span><span class="sxs-lookup"><span data-stu-id="bb311-140">The `getBindingId` method uses the ID of the temporary `Binding` object to retrieve the `Binding` object that raised the event.</span></span>

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

### <a name="cell-format-usestandardheight-and-usestandardwidth-issues"></a><span data-ttu-id="bb311-141">Formato de célula `useStandardHeight` e `useStandardWidth` problemas</span><span class="sxs-lookup"><span data-stu-id="bb311-141">Cell format `useStandardHeight` and `useStandardWidth` issues</span></span>

<span data-ttu-id="bb311-142">A [propriedade useStandardHeight](/javascript/api/excel/excel.cellpropertiesformat#useStandardHeight) `CellPropertiesFormat` não funciona corretamente no Excel na Web.</span><span class="sxs-lookup"><span data-stu-id="bb311-142">The [useStandardHeight](/javascript/api/excel/excel.cellpropertiesformat#useStandardHeight) property of `CellPropertiesFormat` doesn't work properly in Excel on the web.</span></span> <span data-ttu-id="bb311-143">Devido a um problema na interface do usuário Excel na Web, definir a propriedade para calcular a altura `useStandardHeight` `true` imprecisamente nessa plataforma.</span><span class="sxs-lookup"><span data-stu-id="bb311-143">Due to an issue in the Excel on the web UI, setting the `useStandardHeight` property to `true` calculates height imprecisely on this platform.</span></span> <span data-ttu-id="bb311-144">Por exemplo, uma altura padrão **de 14** é modificada para **14,25** em Excel na Web.</span><span class="sxs-lookup"><span data-stu-id="bb311-144">For example, a standard height of **14** is modified to **14.25** in Excel on the web.</span></span>

<span data-ttu-id="bb311-145">Em todas as plataformas, [as propriedades useStandardHeight](/javascript/api/excel/excel.cellpropertiesformat#useStandardHeight) e [useStandardWidth](/javascript/api/excel/excel.cellpropertiesformat#useStandardWidth) devem ser definidas `CellPropertiesFormat` apenas como `true` .</span><span class="sxs-lookup"><span data-stu-id="bb311-145">On all platforms, the [useStandardHeight](/javascript/api/excel/excel.cellpropertiesformat#useStandardHeight) and [useStandardWidth](/javascript/api/excel/excel.cellpropertiesformat#useStandardWidth) properties of `CellPropertiesFormat` are only intended to be set to `true`.</span></span> <span data-ttu-id="bb311-146">Definir essas propriedades como `false` não tem efeito.</span><span class="sxs-lookup"><span data-stu-id="bb311-146">Setting these properties to `false` has no effect.</span></span> 

### <a name="range-getimage-method-unsupported-on-excel-for-mac"></a><span data-ttu-id="bb311-147">Método `getImage` Range sem suporte no Excel para Mac</span><span class="sxs-lookup"><span data-stu-id="bb311-147">Range `getImage` method unsupported on Excel for Mac</span></span>

<span data-ttu-id="bb311-148">O método [Range getImage](/javascript/api/excel/excel.range#getImage__) não tem suporte no Excel para Mac.</span><span class="sxs-lookup"><span data-stu-id="bb311-148">The Range [getImage](/javascript/api/excel/excel.range#getImage__) method isn't currently supported in Excel for Mac.</span></span> <span data-ttu-id="bb311-149">Consulte [OfficeDev/office-js Issue #235](https://github.com/OfficeDev/office-js/issues/235) para o status atual.</span><span class="sxs-lookup"><span data-stu-id="bb311-149">See [OfficeDev/office-js Issue #235](https://github.com/OfficeDev/office-js/issues/235) for the current status.</span></span>

### <a name="range-return-character-limit"></a><span data-ttu-id="bb311-150">Limite de caracteres de retorno de intervalo</span><span class="sxs-lookup"><span data-stu-id="bb311-150">Range return character limit</span></span>

<span data-ttu-id="bb311-151">Os [métodos Worksheet.getRange(address)](/javascript/api/excel/excel.worksheet#getRange_address_) e [Worksheet.getRanges(address)](/javascript/api/excel/excel.worksheet#getRanges_address_) têm um limite de cadeia de caracteres de endereço de 8192 caracteres.</span><span class="sxs-lookup"><span data-stu-id="bb311-151">The [Worksheet.getRange(address)](/javascript/api/excel/excel.worksheet#getRange_address_) and [Worksheet.getRanges(address)](/javascript/api/excel/excel.worksheet#getRanges_address_) methods have an address string limit of 8192 characters.</span></span> <span data-ttu-id="bb311-152">Quando esse limite é excedido, a cadeia de caracteres de endereço é truncada para 8192 caracteres.</span><span class="sxs-lookup"><span data-stu-id="bb311-152">When this limit is exceeded, the address string is truncated to 8192 characters.</span></span>

## <a name="see-also"></a><span data-ttu-id="bb311-153">Confira também</span><span class="sxs-lookup"><span data-stu-id="bb311-153">See also</span></span>

- [<span data-ttu-id="bb311-154">Solucionar erros de desenvolvimento com Office de complementos</span><span class="sxs-lookup"><span data-stu-id="bb311-154">Troubleshoot development errors with Office Add-ins</span></span>](../testing/troubleshoot-development-errors.md)
- [<span data-ttu-id="bb311-155">Solucionar erros de usuários com Suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="bb311-155">Troubleshoot user errors with Office Add-ins</span></span>](../testing/testing-and-troubleshooting.md)
