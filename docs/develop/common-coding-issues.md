---
title: Diretrizes de codificação para problemas comuns e comportamentos de plataforma inesperados
description: Uma lista de problemas da plataforma de API JavaScript do Office frequentemente encontrada pelos desenvolvedores.
ms.date: 07/23/2020
localization_priority: Normal
ms.openlocfilehash: 8f604acaee308c3bd04e181719b091eb948d63ee
ms.sourcegitcommit: 7d5407d3900d2ad1feae79a4bc038afe50568be0
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 07/30/2020
ms.locfileid: "46530454"
---
# <a name="coding-guidance-for-common-issues-and-unexpected-platform-behaviors"></a><span data-ttu-id="bc00f-103">Diretrizes de codificação para problemas comuns e comportamentos de plataforma inesperados</span><span class="sxs-lookup"><span data-stu-id="bc00f-103">Coding guidance for common issues and unexpected platform behaviors</span></span>

<span data-ttu-id="bc00f-104">Este artigo realça aspectos da API JavaScript do Office que podem resultar em comportamento inesperado ou exigir padrões de codificação específicos para obter o resultado desejado.</span><span class="sxs-lookup"><span data-stu-id="bc00f-104">This article highlights aspects of the Office JavaScript API that may result in unexpected behavior or require specific coding patterns to achieve the desired outcome.</span></span> <span data-ttu-id="bc00f-105">Se você encontrar um problema que pertença à lista, informe-nos usando o formulário de comentários na parte inferior do artigo.</span><span class="sxs-lookup"><span data-stu-id="bc00f-105">If you encounter an issue that belongs in this list, please let us know by using the feedback form at the bottom of the article.</span></span>

## <a name="common-apis-and-outlook-apis-are-not-promise-based"></a><span data-ttu-id="bc00f-106">APIs comuns e APIs do Outlook não são baseados em promessa</span><span class="sxs-lookup"><span data-stu-id="bc00f-106">Common APIs and Outlook APIs are not promise-based</span></span>

<span data-ttu-id="bc00f-107">As [APIs comuns](/javascript/api/office) (aquelas que não estão vinculadas a um host específico do Office) e [APIs do Outlook](/javascript/api/outlook) usam um modelo de programação baseado em retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="bc00f-107">The [Common APIs](/javascript/api/office) (those that are not tied to a particular Office host) and [Outlook APIs](/javascript/api/outlook) use a callback-based programming model.</span></span> <span data-ttu-id="bc00f-108">A interação com o documento subjacente do Office requer uma chamada de leitura ou gravação assíncrona que especifica um retorno de chamada a ser executado quando a operação for concluída.</span><span class="sxs-lookup"><span data-stu-id="bc00f-108">Interacting with the underlying Office document requires an asynchronous read or write call that specifies a callback to be ran when the operation completes.</span></span> <span data-ttu-id="bc00f-109">Para obter um exemplo desse padrão, consulte [Document. getFileAsync](/javascript/api/office/office.document#getfileasync-filetype--options--callback-).</span><span class="sxs-lookup"><span data-stu-id="bc00f-109">For an example of this pattern, see [Document.getFileAsync](/javascript/api/office/office.document#getfileasync-filetype--options--callback-).</span></span>

<span data-ttu-id="bc00f-110">Esses métodos comuns de API e API do Outlook não retornam [promessas](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise).</span><span class="sxs-lookup"><span data-stu-id="bc00f-110">These Common API and Outlook API methods do not return [Promises](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise).</span></span> <span data-ttu-id="bc00f-111">Portanto, você não pode usar [Await](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Operators/await) para pausar a execução até que a operação assíncrona seja concluída.</span><span class="sxs-lookup"><span data-stu-id="bc00f-111">Therefore, you cannot use [await](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Operators/await) to pause the execution until the asynchronous operation completes.</span></span> <span data-ttu-id="bc00f-112">Se você precisar `await` de comportamento, você pode encapsule a chamada do método em uma promessa criada explicitamente.</span><span class="sxs-lookup"><span data-stu-id="bc00f-112">If you need `await` behavior, you can wrap the method call in an explicitly created Promise.</span></span>

```js
readDocumentFileAsync(): Promise<any> {
    return new Promise((resolve, reject) => {
        const chunkSize = 65536;
        const self = this;

        Office.context.document.getFileAsync(Office.FileType.Compressed, { sliceSize: chunkSize }, (asyncResult) => {
            if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                reject(asyncResult.error);
            } else {
                // `getAllSlices` is a Promise-wrapped implementation of File.getSliceAsync.
                self.getAllSlices(asyncResult.value).then(result => {
                    if (result.IsSuccess) {
                        resolve(result.Data);
                    } else {
                        reject(asyncResult.error);
                    }
                });
            }
        });
    });
}
```

> [!NOTE]
> <span data-ttu-id="bc00f-113">A documentação de referência contém a implementação com a promessa do [arquivo. getSliceAsync](/javascript/api/office/office.file#getsliceasync-sliceindex--callback-).</span><span class="sxs-lookup"><span data-stu-id="bc00f-113">The reference documentation contains the Promise-wrapped implementation of [File.getSliceAsync](/javascript/api/office/office.file#getsliceasync-sliceindex--callback-).</span></span>

## <a name="some-properties-cannot-be-set-directly"></a><span data-ttu-id="bc00f-114">Algumas propriedades não podem ser definidas diretamente</span><span class="sxs-lookup"><span data-stu-id="bc00f-114">Some properties cannot be set directly</span></span>

> [!NOTE]
> <span data-ttu-id="bc00f-115">Esta seção só se aplica às APIs específicas do host para Excel e Word.</span><span class="sxs-lookup"><span data-stu-id="bc00f-115">This section only applies to the host-specific APIs for Excel and Word.</span></span>

<span data-ttu-id="bc00f-116">Algumas propriedades não podem ser definidas, apesar de serem graváveis.</span><span class="sxs-lookup"><span data-stu-id="bc00f-116">Some properties cannot be set, despite being writable.</span></span> <span data-ttu-id="bc00f-117">Essas propriedades fazem parte de uma propriedade pai que deve ser definida como um único objeto.</span><span class="sxs-lookup"><span data-stu-id="bc00f-117">These properties are part of a parent property that must be set as a single object.</span></span> <span data-ttu-id="bc00f-118">Isso ocorre porque a propriedade Parent depende das subpropriedades que têm relações lógicas específicas.</span><span class="sxs-lookup"><span data-stu-id="bc00f-118">This is because that parent property relies on the subproperties having specific, logical relationships.</span></span> <span data-ttu-id="bc00f-119">Essas propriedades pai devem ser definidas usando a notação literal de objeto para definir o objeto inteiro, em vez de definir as subpropriedades individuais desse objeto.</span><span class="sxs-lookup"><span data-stu-id="bc00f-119">These parent properties must be set using object literal notation to set the entire object, instead of setting that object's individual subproperties.</span></span> <span data-ttu-id="bc00f-120">Um exemplo disso é encontrado no [PageLayout](/javascript/api/excel/excel.pagelayout).</span><span class="sxs-lookup"><span data-stu-id="bc00f-120">One example of this is found in [PageLayout](/javascript/api/excel/excel.pagelayout).</span></span> <span data-ttu-id="bc00f-121">A `zoom` propriedade deve ser definida com um único objeto [PageLayoutZoomOptions](/javascript/api/excel/excel.pagelayoutzoomoptions) , conforme mostrado aqui:</span><span class="sxs-lookup"><span data-stu-id="bc00f-121">The `zoom` property must be set with a single [PageLayoutZoomOptions](/javascript/api/excel/excel.pagelayoutzoomoptions) object, as shown here:</span></span>

```js
// PageLayout.zoom.scale must be set by assigning PageLayout.zoom to a PageLayoutZoomOptions object.
sheet.pageLayout.zoom = { scale: 200 };
```

<span data-ttu-id="bc00f-122">No exemplo anterior, você ***não*** poderá atribuir `zoom` um valor diretamente: `sheet.pageLayout.zoom.scale = 200;` .</span><span class="sxs-lookup"><span data-stu-id="bc00f-122">In the previous example, you would ***not*** be able to directly assign `zoom` a value: `sheet.pageLayout.zoom.scale = 200;`.</span></span> <span data-ttu-id="bc00f-123">Essa instrução gera um erro porque `zoom` não está carregada.</span><span class="sxs-lookup"><span data-stu-id="bc00f-123">That statement throws an error because `zoom` is not loaded.</span></span> <span data-ttu-id="bc00f-124">Mesmo que `zoom` fosse carregado, o conjunto de escala não terá efeito.</span><span class="sxs-lookup"><span data-stu-id="bc00f-124">Even if `zoom` were to be loaded, the set of scale will not take effect.</span></span> <span data-ttu-id="bc00f-125">Todas as operações de contexto acontecem em `zoom` , atualizando o objeto de proxy no suplemento e substituindo os valores definidos localmente.</span><span class="sxs-lookup"><span data-stu-id="bc00f-125">All context operations happen on `zoom`, refreshing the proxy object in the add-in and overwriting locally set values.</span></span>

<span data-ttu-id="bc00f-126">Esse comportamento difere das [Propriedades de navegação](../excel/excel-add-ins-advanced-concepts.md#scalar-and-navigation-properties) , como [Range. Format](/javascript/api/excel/excel.range#format).</span><span class="sxs-lookup"><span data-stu-id="bc00f-126">This behavior differs from [navigational properties](../excel/excel-add-ins-advanced-concepts.md#scalar-and-navigation-properties) like [Range.format](/javascript/api/excel/excel.range#format).</span></span> <span data-ttu-id="bc00f-127">As propriedades de `format` podem ser definidas usando a navegação de objeto, conforme mostrado aqui:</span><span class="sxs-lookup"><span data-stu-id="bc00f-127">Properties of `format` can be set using object navigation, as shown here:</span></span>

```js
// This will set the font size on the range during the next `content.sync()`.
range.format.font.size = 10;
```

<span data-ttu-id="bc00f-128">Você pode identificar uma propriedade que não pode ter suas subpropriedades definidas diretamente verificando seu modificador somente leitura.</span><span class="sxs-lookup"><span data-stu-id="bc00f-128">You can identify a property that cannot have its subproperties directly set by checking its read-only modifier.</span></span> <span data-ttu-id="bc00f-129">Todas as propriedades somente leitura podem ter suas subpropriedades não somente leitura definidas diretamente.</span><span class="sxs-lookup"><span data-stu-id="bc00f-129">All read-only properties can have their non-read-only subproperties directly set.</span></span> <span data-ttu-id="bc00f-130">Propriedades graváveis como `PageLayout.zoom` devem ser definidas com um objeto nesse nível.</span><span class="sxs-lookup"><span data-stu-id="bc00f-130">Writeable properties like `PageLayout.zoom` must be set with an object at that level.</span></span> <span data-ttu-id="bc00f-131">Em Resumo:</span><span class="sxs-lookup"><span data-stu-id="bc00f-131">In summary:</span></span>

- <span data-ttu-id="bc00f-132">Propriedade somente leitura: as subpropriedades podem ser definidas por meio de navegação.</span><span class="sxs-lookup"><span data-stu-id="bc00f-132">Read-only property: Subproperties can be set through navigation.</span></span>
- <span data-ttu-id="bc00f-133">Propriedade writable: as subpropriedades não podem ser definidas por meio de navegação (devem ser definidas como parte da atribuição de objeto pai inicial).</span><span class="sxs-lookup"><span data-stu-id="bc00f-133">Writable property: Subproperties cannot be set through navigation (must be set as part of the initial parent object assignment).</span></span>

## <a name="setting-read-only-properties"></a><span data-ttu-id="bc00f-134">Configuração de propriedades somente leitura</span><span class="sxs-lookup"><span data-stu-id="bc00f-134">Setting read-only properties</span></span>

<span data-ttu-id="bc00f-135">As [definições do TypeScript](referencing-the-javascript-api-for-office-library-from-its-cdn.md) para o Office js especificam quais propriedades de objeto são somente leitura.</span><span class="sxs-lookup"><span data-stu-id="bc00f-135">The [TypeScript definitions](referencing-the-javascript-api-for-office-library-from-its-cdn.md) for Office JS specify which object properties are read-only.</span></span> <span data-ttu-id="bc00f-136">Se você tentar definir uma propriedade somente leitura, a operação de gravação falhará silenciosamente, sem nenhum erro.</span><span class="sxs-lookup"><span data-stu-id="bc00f-136">If you attempt to set a read-only property, the write operation will fail silently, with no error thrown.</span></span> <span data-ttu-id="bc00f-137">O exemplo a seguir tenta erroneamente definir a propriedade somente leitura [Chart.ID](/javascript/api/excel/excel.chart#id).</span><span class="sxs-lookup"><span data-stu-id="bc00f-137">The following example erroneously attempts to set the read-only property [Chart.id](/javascript/api/excel/excel.chart#id).</span></span>

```js
// This will do nothing, since `id` is a read-only property.
myChart.id = "5";
```

## <a name="removing-event-handlers"></a><span data-ttu-id="bc00f-138">Remover manipuladores de eventos</span><span class="sxs-lookup"><span data-stu-id="bc00f-138">Removing event handlers</span></span>

<span data-ttu-id="bc00f-139">Manipuladores de eventos devem ser removidos usando o mesmo `RequestContext` em que foram adicionados.</span><span class="sxs-lookup"><span data-stu-id="bc00f-139">Event handlers must be removed using the same `RequestContext` in which they were added.</span></span> <span data-ttu-id="bc00f-140">Se você precisar que seu suplemento remova um manipulador de eventos durante a execução, será necessário armazenar o objeto Context usado para adicionar o manipulador.</span><span class="sxs-lookup"><span data-stu-id="bc00f-140">If you need your add-in to remove an event handler while running, you'll need to store the context object used to add the handler.</span></span>

```js
Excel.run(async (context) => {
    [...]

    // To later remove an event handler, store the context somewhere accessible to the handler removal function.
    // You may find it helpful to also store the event handler object and associate it with the context.
    selectionChangedHandler = myWorksheet.onSelectionChanged.add(callback);
    savedContext = currentContext;
    return context.sync();
}
```

## <a name="supporting-internet-explorer"></a><span data-ttu-id="bc00f-141">Suporte do Internet Explorer</span><span class="sxs-lookup"><span data-stu-id="bc00f-141">Supporting Internet Explorer</span></span>

[!INCLUDE [How to support IE](../includes/es5-support.md)]

## <a name="excel-specific-issues"></a><span data-ttu-id="bc00f-142">Problemas específicos do Excel</span><span class="sxs-lookup"><span data-stu-id="bc00f-142">Excel-specific issues</span></span>

### <a name="excel-data-transfer-limits"></a><span data-ttu-id="bc00f-143">Limites de transferência de dados do Excel</span><span class="sxs-lookup"><span data-stu-id="bc00f-143">Excel data transfer limits</span></span>

<span data-ttu-id="bc00f-144">Se você estiver criando um suplemento do Excel, esteja ciente das seguintes limitações de tamanho ao interagir com a pasta de trabalho:</span><span class="sxs-lookup"><span data-stu-id="bc00f-144">If you're building an Excel add-in, be aware of the following size limitations when interacting with the workbook:</span></span>

- <span data-ttu-id="bc00f-145">O Excel na Web tem um limite de tamanho de conteúdo para solicitações e respostas de 5 MB.</span><span class="sxs-lookup"><span data-stu-id="bc00f-145">Excel on the web has a payload size limit for requests and responses of 5MB.</span></span> <span data-ttu-id="bc00f-146">`RichAPI.Error` será lançado se esse limite for excedido.</span><span class="sxs-lookup"><span data-stu-id="bc00f-146">`RichAPI.Error` will be thrown if that limit is exceeded.</span></span>
- <span data-ttu-id="bc00f-147">Um intervalo está limitado a 5 milhões células para operações Get.</span><span class="sxs-lookup"><span data-stu-id="bc00f-147">A range is limited to five million cells for get operations.</span></span>

<span data-ttu-id="bc00f-148">Se você espera que a entrada do usuário exceda esses limites, verifique os dados antes de chamar `context.sync()` .</span><span class="sxs-lookup"><span data-stu-id="bc00f-148">If you expect user input to exceed these limits, be sure to check the data before calling `context.sync()`.</span></span> <span data-ttu-id="bc00f-149">Divida a operação em partes menores, conforme necessário.</span><span class="sxs-lookup"><span data-stu-id="bc00f-149">Split the operation into smaller pieces as needed.</span></span> <span data-ttu-id="bc00f-150">Certifique-se de chamar `context.sync()` para cada suboperação para evitar que as operações sejam encaixadas novamente.</span><span class="sxs-lookup"><span data-stu-id="bc00f-150">Be sure to call `context.sync()` for each sub-operation to avoid those operations getting batched together again.</span></span>

<span data-ttu-id="bc00f-151">Essas limitações são normalmente excedidos por intervalos grandes.</span><span class="sxs-lookup"><span data-stu-id="bc00f-151">These limitations are typically exceeded by large ranges.</span></span> <span data-ttu-id="bc00f-152">O suplemento pode ser capaz de usar o [RangeAreas](/javascript/api/excel/excel.rangeareas) para atualizar as células estrategicamente em um intervalo maior.</span><span class="sxs-lookup"><span data-stu-id="bc00f-152">Your add-in might be able to use [RangeAreas](/javascript/api/excel/excel.rangeareas) to strategically update cells within a larger range.</span></span> <span data-ttu-id="bc00f-153">Confira [trabalhar com vários intervalos simultaneamente em suplementos do Excel](../excel/excel-add-ins-multiple-ranges.md) para obter mais informações.</span><span class="sxs-lookup"><span data-stu-id="bc00f-153">See [Work with multiple ranges simultaneously in Excel add-ins](../excel/excel-add-ins-multiple-ranges.md) for more information.</span></span>

### <a name="api-limitations-when-the-active-workbook-switches"></a><span data-ttu-id="bc00f-154">Limitações de API quando a pasta de trabalho ativa alterna</span><span class="sxs-lookup"><span data-stu-id="bc00f-154">API limitations when the active workbook switches</span></span>

<span data-ttu-id="bc00f-155">Os suplementos para Excel se destinam a operar em uma única pasta de trabalho por vez.</span><span class="sxs-lookup"><span data-stu-id="bc00f-155">Add-ins for Excel are intended to operate on a single workbook at a time.</span></span> <span data-ttu-id="bc00f-156">Os erros podem ocorrer quando uma pasta de trabalho separada da que está executando o suplemento Obtém o foco.</span><span class="sxs-lookup"><span data-stu-id="bc00f-156">Errors can arise when a workbook that is separate from the one running the add-in gains focus.</span></span> <span data-ttu-id="bc00f-157">Isso ocorre apenas quando determinados métodos estão no processo de chamada quando o foco é alterado.</span><span class="sxs-lookup"><span data-stu-id="bc00f-157">This only happens when particular methods are in the process of being called when the focus changes.</span></span>

<span data-ttu-id="bc00f-158">As seguintes APIs são afetadas por essa opção de pasta de trabalho:</span><span class="sxs-lookup"><span data-stu-id="bc00f-158">The following APIs are affected by this workbook switch:</span></span>

|<span data-ttu-id="bc00f-159">API JavaScript do Excel</span><span class="sxs-lookup"><span data-stu-id="bc00f-159">Excel JavaScript API</span></span> | <span data-ttu-id="bc00f-160">Erro gerado</span><span class="sxs-lookup"><span data-stu-id="bc00f-160">Error thrown</span></span> |
|--|--|
| `Chart.activate` | <span data-ttu-id="bc00f-161">GeneralException</span><span class="sxs-lookup"><span data-stu-id="bc00f-161">GeneralException</span></span> |
| `Range.select` | <span data-ttu-id="bc00f-162">GeneralException</span><span class="sxs-lookup"><span data-stu-id="bc00f-162">GeneralException</span></span> |
| `Table.clearFilters` | <span data-ttu-id="bc00f-163">GeneralException</span><span class="sxs-lookup"><span data-stu-id="bc00f-163">GeneralException</span></span> |
| `Workbook.getActiveCell`  | <span data-ttu-id="bc00f-164">InvalidSelection</span><span class="sxs-lookup"><span data-stu-id="bc00f-164">InvalidSelection</span></span>|
| `Workbook.getSelectedRange` | <span data-ttu-id="bc00f-165">InvalidSelection</span><span class="sxs-lookup"><span data-stu-id="bc00f-165">InvalidSelection</span></span>|
| `Workbook.getSelectedRanges`  | <span data-ttu-id="bc00f-166">InvalidSelection</span><span class="sxs-lookup"><span data-stu-id="bc00f-166">InvalidSelection</span></span>|
| `Worksheet.activate` | <span data-ttu-id="bc00f-167">GeneralException</span><span class="sxs-lookup"><span data-stu-id="bc00f-167">GeneralException</span></span> |
| `Worksheet.delete`  | <span data-ttu-id="bc00f-168">InvalidSelection</span><span class="sxs-lookup"><span data-stu-id="bc00f-168">InvalidSelection</span></span>|
| `Worksheet.gridlines` | <span data-ttu-id="bc00f-169">GeneralException</span><span class="sxs-lookup"><span data-stu-id="bc00f-169">GeneralException</span></span> |
| `Worksheet.showHeadings` | <span data-ttu-id="bc00f-170">GeneralException</span><span class="sxs-lookup"><span data-stu-id="bc00f-170">GeneralException</span></span> |
| `WorksheetCollection.add` | <span data-ttu-id="bc00f-171">GeneralException</span><span class="sxs-lookup"><span data-stu-id="bc00f-171">GeneralException</span></span> |
| `WorksheetFreezePanes.freezeAt` | <span data-ttu-id="bc00f-172">GeneralException</span><span class="sxs-lookup"><span data-stu-id="bc00f-172">GeneralException</span></span> |
| `WorksheetFreezePanes.freezeColumns` | <span data-ttu-id="bc00f-173">GeneralException</span><span class="sxs-lookup"><span data-stu-id="bc00f-173">GeneralException</span></span> |
| `WorksheetFreezePanes.freezeRows` | <span data-ttu-id="bc00f-174">GeneralException</span><span class="sxs-lookup"><span data-stu-id="bc00f-174">GeneralException</span></span> |
| `WorksheetFreezePanes.getLocationOrNullObject`| <span data-ttu-id="bc00f-175">GeneralException</span><span class="sxs-lookup"><span data-stu-id="bc00f-175">GeneralException</span></span> |
| `WorksheetFreezePanes.unfreeze` | <span data-ttu-id="bc00f-176">GeneralException</span><span class="sxs-lookup"><span data-stu-id="bc00f-176">GeneralException</span></span> |

> [!NOTE]
> <span data-ttu-id="bc00f-177">Isso aplica-se apenas a várias pastas de trabalho do Excel abertas no Windows ou Mac.</span><span class="sxs-lookup"><span data-stu-id="bc00f-177">This only applies to multiple Excel workbooks open on Windows or Mac.</span></span>

### <a name="coauthoring"></a><span data-ttu-id="bc00f-178">Coautoria</span><span class="sxs-lookup"><span data-stu-id="bc00f-178">Coauthoring</span></span>

<span data-ttu-id="bc00f-179">Veja [coautoria em suplementos do Excel](../excel/co-authoring-in-excel-add-ins.md) para padrões a serem usados com eventos em um ambiente de coautoria.</span><span class="sxs-lookup"><span data-stu-id="bc00f-179">See [Coauthoring in Excel add-ins](../excel/co-authoring-in-excel-add-ins.md) for patterns to use with events in a coauthoring environment.</span></span> <span data-ttu-id="bc00f-180">O artigo também aborda possíveis conflitos de mesclagem ao usar determinadas APIs, como [`TableRowCollection.add`](/javascript/api/excel/excel.tablerowcollection#add-index--values-) .</span><span class="sxs-lookup"><span data-stu-id="bc00f-180">The article also discusses potential merge conflicts when using certain APIs, such as [`TableRowCollection.add`](/javascript/api/excel/excel.tablerowcollection#add-index--values-).</span></span>

## <a name="see-also"></a><span data-ttu-id="bc00f-181">Confira também</span><span class="sxs-lookup"><span data-stu-id="bc00f-181">See also</span></span>

- <span data-ttu-id="bc00f-182">[OfficeDev/Office-js](https://github.com/OfficeDev/office-js/issues): o local para relatar e exibir problemas com a plataforma de suplementos do Office e APIs JavaScript.</span><span class="sxs-lookup"><span data-stu-id="bc00f-182">[OfficeDev/office-js](https://github.com/OfficeDev/office-js/issues): The place to report and view issues with the Office Add-ins platform and JavaScript APIs.</span></span>
- <span data-ttu-id="bc00f-183">[Estouro de pilha](https://stackoverflow.com/questions/tagged/office-js): o local para solicitar e exibir perguntas de programação sobre as APIs JavaScript do Office.</span><span class="sxs-lookup"><span data-stu-id="bc00f-183">[Stack Overflow](https://stackoverflow.com/questions/tagged/office-js): The place to ask and view programming questions about the Office JavaScript APIs.</span></span> <span data-ttu-id="bc00f-184">Certifique-se de aplicar a marca "Office-js" à sua pergunta ao postar no estouro de pilha.</span><span class="sxs-lookup"><span data-stu-id="bc00f-184">Be sure to apply the "office-js" tag to your question when posting to Stack Overflow.</span></span>
- <span data-ttu-id="bc00f-185">[UserVoice](https://officespdev.uservoice.com/): o local para sugerir novos recursos para a plataforma de suplementos do Office e APIs JavaScript do Office.</span><span class="sxs-lookup"><span data-stu-id="bc00f-185">[UserVoice](https://officespdev.uservoice.com/): The place to suggest new features for the Office Add-ins platform and Office JavaScript APIs.</span></span>
