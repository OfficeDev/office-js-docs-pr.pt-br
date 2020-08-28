---
title: Diretrizes de codificação para problemas comuns e comportamentos de plataforma inesperados
description: Uma lista de problemas da plataforma de API JavaScript do Office frequentemente encontrada pelos desenvolvedores.
ms.date: 07/29/2020
localization_priority: Normal
ms.openlocfilehash: f6d6a31059b32550e3176ed278d7da4c2c7a6c68
ms.sourcegitcommit: 9609bd5b4982cdaa2ea7637709a78a45835ffb19
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 08/28/2020
ms.locfileid: "47292908"
---
# <a name="coding-guidance-for-common-issues-and-unexpected-platform-behaviors"></a><span data-ttu-id="e6240-103">Diretrizes de codificação para problemas comuns e comportamentos de plataforma inesperados</span><span class="sxs-lookup"><span data-stu-id="e6240-103">Coding guidance for common issues and unexpected platform behaviors</span></span>

<span data-ttu-id="e6240-104">Este artigo realça aspectos da API JavaScript do Office que podem resultar em comportamento inesperado ou exigir padrões de codificação específicos para obter o resultado desejado.</span><span class="sxs-lookup"><span data-stu-id="e6240-104">This article highlights aspects of the Office JavaScript API that may result in unexpected behavior or require specific coding patterns to achieve the desired outcome.</span></span> <span data-ttu-id="e6240-105">Se você encontrar um problema que pertença à lista, informe-nos usando o formulário de comentários na parte inferior do artigo.</span><span class="sxs-lookup"><span data-stu-id="e6240-105">If you encounter an issue that belongs in this list, please let us know by using the feedback form at the bottom of the article.</span></span>

## <a name="common-apis-and-outlook-apis-are-not-promise-based"></a><span data-ttu-id="e6240-106">APIs comuns e APIs do Outlook não são baseados em promessa</span><span class="sxs-lookup"><span data-stu-id="e6240-106">Common APIs and Outlook APIs are not promise-based</span></span>

<span data-ttu-id="e6240-107">As [APIs comuns](/javascript/api/office) (aquelas que não estão vinculadas a um aplicativo específico do Office) e [APIs do Outlook](/javascript/api/outlook) usam um modelo de programação baseado em retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="e6240-107">The [Common APIs](/javascript/api/office) (those that are not tied to a particular Office application) and [Outlook APIs](/javascript/api/outlook) use a callback-based programming model.</span></span> <span data-ttu-id="e6240-108">A interação com o documento subjacente do Office requer uma chamada de leitura ou gravação assíncrona que especifica um retorno de chamada a ser executado quando a operação for concluída.</span><span class="sxs-lookup"><span data-stu-id="e6240-108">Interacting with the underlying Office document requires an asynchronous read or write call that specifies a callback to be run when the operation completes.</span></span> <span data-ttu-id="e6240-109">Para obter um exemplo desse padrão, consulte [Document. getFileAsync](/javascript/api/office/office.document#getfileasync-filetype--options--callback-).</span><span class="sxs-lookup"><span data-stu-id="e6240-109">For an example of this pattern, see [Document.getFileAsync](/javascript/api/office/office.document#getfileasync-filetype--options--callback-).</span></span>

<span data-ttu-id="e6240-110">Esses métodos comuns de API e API do Outlook não retornam [promessas](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise).</span><span class="sxs-lookup"><span data-stu-id="e6240-110">These Common API and Outlook API methods do not return [Promises](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise).</span></span> <span data-ttu-id="e6240-111">Portanto, você não pode usar [Await](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Operators/await) para pausar a execução até que a operação assíncrona seja concluída.</span><span class="sxs-lookup"><span data-stu-id="e6240-111">Therefore, you cannot use [await](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Operators/await) to pause the execution until the asynchronous operation completes.</span></span> <span data-ttu-id="e6240-112">Se você precisar `await` de comportamento, você pode encapsule a chamada do método em uma promessa criada explicitamente.</span><span class="sxs-lookup"><span data-stu-id="e6240-112">If you need `await` behavior, you can wrap the method call in an explicitly created Promise.</span></span>

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
> <span data-ttu-id="e6240-113">A documentação de referência contém a implementação com a promessa do [arquivo. getSliceAsync](/javascript/api/office/office.file#getsliceasync-sliceindex--callback-).</span><span class="sxs-lookup"><span data-stu-id="e6240-113">The reference documentation contains the Promise-wrapped implementation of [File.getSliceAsync](/javascript/api/office/office.file#getsliceasync-sliceindex--callback-).</span></span>

## <a name="some-properties-cannot-be-set-directly"></a><span data-ttu-id="e6240-114">Algumas propriedades não podem ser definidas diretamente</span><span class="sxs-lookup"><span data-stu-id="e6240-114">Some properties cannot be set directly</span></span>

> [!NOTE]
> <span data-ttu-id="e6240-115">Esta seção só se aplica às APIs específicas do aplicativo para Excel e Word.</span><span class="sxs-lookup"><span data-stu-id="e6240-115">This section only applies to the application-specific APIs for Excel and Word.</span></span>

<span data-ttu-id="e6240-116">Algumas propriedades não podem ser definidas, apesar de serem graváveis.</span><span class="sxs-lookup"><span data-stu-id="e6240-116">Some properties cannot be set, despite being writable.</span></span> <span data-ttu-id="e6240-117">Essas propriedades fazem parte de uma propriedade pai que deve ser definida como um único objeto.</span><span class="sxs-lookup"><span data-stu-id="e6240-117">These properties are part of a parent property that must be set as a single object.</span></span> <span data-ttu-id="e6240-118">Isso ocorre porque a propriedade Parent depende das subpropriedades que têm relações lógicas específicas.</span><span class="sxs-lookup"><span data-stu-id="e6240-118">This is because that parent property relies on the subproperties having specific, logical relationships.</span></span> <span data-ttu-id="e6240-119">Essas propriedades pai devem ser definidas usando a notação literal de objeto para definir o objeto inteiro, em vez de definir as subpropriedades individuais desse objeto.</span><span class="sxs-lookup"><span data-stu-id="e6240-119">These parent properties must be set using object literal notation to set the entire object, instead of setting that object's individual subproperties.</span></span> <span data-ttu-id="e6240-120">Um exemplo disso é encontrado no [PageLayout](/javascript/api/excel/excel.pagelayout).</span><span class="sxs-lookup"><span data-stu-id="e6240-120">One example of this is found in [PageLayout](/javascript/api/excel/excel.pagelayout).</span></span> <span data-ttu-id="e6240-121">A `zoom` propriedade deve ser definida com um único objeto [PageLayoutZoomOptions](/javascript/api/excel/excel.pagelayoutzoomoptions) , conforme mostrado aqui:</span><span class="sxs-lookup"><span data-stu-id="e6240-121">The `zoom` property must be set with a single [PageLayoutZoomOptions](/javascript/api/excel/excel.pagelayoutzoomoptions) object, as shown here:</span></span>

```js
// PageLayout.zoom.scale must be set by assigning PageLayout.zoom to a PageLayoutZoomOptions object.
sheet.pageLayout.zoom = { scale: 200 };
```

<span data-ttu-id="e6240-122">No exemplo anterior, você ***não*** poderá atribuir `zoom` um valor diretamente: `sheet.pageLayout.zoom.scale = 200;` .</span><span class="sxs-lookup"><span data-stu-id="e6240-122">In the previous example, you would ***not*** be able to directly assign `zoom` a value: `sheet.pageLayout.zoom.scale = 200;`.</span></span> <span data-ttu-id="e6240-123">Essa instrução gera um erro porque `zoom` não está carregada.</span><span class="sxs-lookup"><span data-stu-id="e6240-123">That statement throws an error because `zoom` is not loaded.</span></span> <span data-ttu-id="e6240-124">Mesmo que `zoom` fosse carregado, o conjunto de escala não terá efeito.</span><span class="sxs-lookup"><span data-stu-id="e6240-124">Even if `zoom` were to be loaded, the set of scale will not take effect.</span></span> <span data-ttu-id="e6240-125">Todas as operações de contexto acontecem em `zoom` , atualizando o objeto de proxy no suplemento e substituindo os valores definidos localmente.</span><span class="sxs-lookup"><span data-stu-id="e6240-125">All context operations happen on `zoom`, refreshing the proxy object in the add-in and overwriting locally set values.</span></span>

<span data-ttu-id="e6240-126">Esse comportamento difere das [Propriedades de navegação](application-specific-api-model.md#scalar-and-navigation-properties) , como [Range. Format](/javascript/api/excel/excel.range#format).</span><span class="sxs-lookup"><span data-stu-id="e6240-126">This behavior differs from [navigational properties](application-specific-api-model.md#scalar-and-navigation-properties) like [Range.format](/javascript/api/excel/excel.range#format).</span></span> <span data-ttu-id="e6240-127">As propriedades de `format` podem ser definidas usando a navegação de objeto, conforme mostrado aqui:</span><span class="sxs-lookup"><span data-stu-id="e6240-127">Properties of `format` can be set using object navigation, as shown here:</span></span>

```js
// This will set the font size on the range during the next `content.sync()`.
range.format.font.size = 10;
```

<span data-ttu-id="e6240-128">Você pode identificar uma propriedade que não pode ter suas subpropriedades definidas diretamente verificando seu modificador somente leitura.</span><span class="sxs-lookup"><span data-stu-id="e6240-128">You can identify a property that cannot have its subproperties directly set by checking its read-only modifier.</span></span> <span data-ttu-id="e6240-129">Todas as propriedades somente leitura podem ter suas subpropriedades não somente leitura definidas diretamente.</span><span class="sxs-lookup"><span data-stu-id="e6240-129">All read-only properties can have their non-read-only subproperties directly set.</span></span> <span data-ttu-id="e6240-130">Propriedades graváveis como `PageLayout.zoom` devem ser definidas com um objeto nesse nível.</span><span class="sxs-lookup"><span data-stu-id="e6240-130">Writeable properties like `PageLayout.zoom` must be set with an object at that level.</span></span> <span data-ttu-id="e6240-131">Em Resumo:</span><span class="sxs-lookup"><span data-stu-id="e6240-131">In summary:</span></span>

- <span data-ttu-id="e6240-132">Propriedade somente leitura: as subpropriedades podem ser definidas por meio de navegação.</span><span class="sxs-lookup"><span data-stu-id="e6240-132">Read-only property: Subproperties can be set through navigation.</span></span>
- <span data-ttu-id="e6240-133">Propriedade writable: as subpropriedades não podem ser definidas por meio de navegação (devem ser definidas como parte da atribuição de objeto pai inicial).</span><span class="sxs-lookup"><span data-stu-id="e6240-133">Writable property: Subproperties cannot be set through navigation (must be set as part of the initial parent object assignment).</span></span>

## <a name="setting-read-only-properties"></a><span data-ttu-id="e6240-134">Configuração de propriedades somente leitura</span><span class="sxs-lookup"><span data-stu-id="e6240-134">Setting read-only properties</span></span>

<span data-ttu-id="e6240-135">As [definições do TypeScript](referencing-the-javascript-api-for-office-library-from-its-cdn.md) para o Office js especificam quais propriedades de objeto são somente leitura.</span><span class="sxs-lookup"><span data-stu-id="e6240-135">The [TypeScript definitions](referencing-the-javascript-api-for-office-library-from-its-cdn.md) for Office JS specify which object properties are read-only.</span></span> <span data-ttu-id="e6240-136">Se você tentar definir uma propriedade somente leitura, a operação de gravação falhará silenciosamente, sem nenhum erro.</span><span class="sxs-lookup"><span data-stu-id="e6240-136">If you attempt to set a read-only property, the write operation will fail silently, with no error thrown.</span></span> <span data-ttu-id="e6240-137">O exemplo a seguir tenta erroneamente definir a propriedade somente leitura [Chart.ID](/javascript/api/excel/excel.chart#id).</span><span class="sxs-lookup"><span data-stu-id="e6240-137">The following example erroneously attempts to set the read-only property [Chart.id](/javascript/api/excel/excel.chart#id).</span></span>

```js
// This will do nothing, since `id` is a read-only property.
myChart.id = "5";
```

## <a name="removing-event-handlers"></a><span data-ttu-id="e6240-138">Remover manipuladores de eventos</span><span class="sxs-lookup"><span data-stu-id="e6240-138">Removing event handlers</span></span>

<span data-ttu-id="e6240-139">Manipuladores de eventos devem ser removidos usando o mesmo `RequestContext` em que foram adicionados.</span><span class="sxs-lookup"><span data-stu-id="e6240-139">Event handlers must be removed using the same `RequestContext` in which they were added.</span></span> <span data-ttu-id="e6240-140">Se você precisar que seu suplemento remova um manipulador de eventos durante a execução, será necessário armazenar o objeto Context usado para adicionar o manipulador.</span><span class="sxs-lookup"><span data-stu-id="e6240-140">If you need your add-in to remove an event handler while running, you'll need to store the context object used to add the handler.</span></span>

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

## <a name="supporting-internet-explorer"></a><span data-ttu-id="e6240-141">Suporte do Internet Explorer</span><span class="sxs-lookup"><span data-stu-id="e6240-141">Supporting Internet Explorer</span></span>

[!INCLUDE [How to support IE](../includes/es5-support.md)]

## <a name="excel-specific-issues"></a><span data-ttu-id="e6240-142">Problemas específicos do Excel</span><span class="sxs-lookup"><span data-stu-id="e6240-142">Excel-specific issues</span></span>

### <a name="api-limitations-when-the-active-workbook-switches"></a><span data-ttu-id="e6240-143">Limitações de API quando a pasta de trabalho ativa alterna</span><span class="sxs-lookup"><span data-stu-id="e6240-143">API limitations when the active workbook switches</span></span>

<span data-ttu-id="e6240-144">Os suplementos para Excel se destinam a operar em uma única pasta de trabalho por vez.</span><span class="sxs-lookup"><span data-stu-id="e6240-144">Add-ins for Excel are intended to operate on a single workbook at a time.</span></span> <span data-ttu-id="e6240-145">Os erros podem ocorrer quando uma pasta de trabalho separada da que está executando o suplemento Obtém o foco.</span><span class="sxs-lookup"><span data-stu-id="e6240-145">Errors can arise when a workbook that is separate from the one running the add-in gains focus.</span></span> <span data-ttu-id="e6240-146">Isso ocorre apenas quando determinados métodos estão no processo de chamada quando o foco é alterado.</span><span class="sxs-lookup"><span data-stu-id="e6240-146">This only happens when particular methods are in the process of being called when the focus changes.</span></span>

<span data-ttu-id="e6240-147">As seguintes APIs são afetadas por essa opção de pasta de trabalho:</span><span class="sxs-lookup"><span data-stu-id="e6240-147">The following APIs are affected by this workbook switch:</span></span>

|<span data-ttu-id="e6240-148">API JavaScript do Excel</span><span class="sxs-lookup"><span data-stu-id="e6240-148">Excel JavaScript API</span></span> | <span data-ttu-id="e6240-149">Erro gerado</span><span class="sxs-lookup"><span data-stu-id="e6240-149">Error thrown</span></span> |
|--|--|
| `Chart.activate` | <span data-ttu-id="e6240-150">GeneralException</span><span class="sxs-lookup"><span data-stu-id="e6240-150">GeneralException</span></span> |
| `Range.select` | <span data-ttu-id="e6240-151">GeneralException</span><span class="sxs-lookup"><span data-stu-id="e6240-151">GeneralException</span></span> |
| `Table.clearFilters` | <span data-ttu-id="e6240-152">GeneralException</span><span class="sxs-lookup"><span data-stu-id="e6240-152">GeneralException</span></span> |
| `Workbook.getActiveCell`  | <span data-ttu-id="e6240-153">InvalidSelection</span><span class="sxs-lookup"><span data-stu-id="e6240-153">InvalidSelection</span></span>|
| `Workbook.getSelectedRange` | <span data-ttu-id="e6240-154">InvalidSelection</span><span class="sxs-lookup"><span data-stu-id="e6240-154">InvalidSelection</span></span>|
| `Workbook.getSelectedRanges`  | <span data-ttu-id="e6240-155">InvalidSelection</span><span class="sxs-lookup"><span data-stu-id="e6240-155">InvalidSelection</span></span>|
| `Worksheet.activate` | <span data-ttu-id="e6240-156">GeneralException</span><span class="sxs-lookup"><span data-stu-id="e6240-156">GeneralException</span></span> |
| `Worksheet.delete`  | <span data-ttu-id="e6240-157">InvalidSelection</span><span class="sxs-lookup"><span data-stu-id="e6240-157">InvalidSelection</span></span>|
| `Worksheet.gridlines` | <span data-ttu-id="e6240-158">GeneralException</span><span class="sxs-lookup"><span data-stu-id="e6240-158">GeneralException</span></span> |
| `Worksheet.showHeadings` | <span data-ttu-id="e6240-159">GeneralException</span><span class="sxs-lookup"><span data-stu-id="e6240-159">GeneralException</span></span> |
| `WorksheetCollection.add` | <span data-ttu-id="e6240-160">GeneralException</span><span class="sxs-lookup"><span data-stu-id="e6240-160">GeneralException</span></span> |
| `WorksheetFreezePanes.freezeAt` | <span data-ttu-id="e6240-161">GeneralException</span><span class="sxs-lookup"><span data-stu-id="e6240-161">GeneralException</span></span> |
| `WorksheetFreezePanes.freezeColumns` | <span data-ttu-id="e6240-162">GeneralException</span><span class="sxs-lookup"><span data-stu-id="e6240-162">GeneralException</span></span> |
| `WorksheetFreezePanes.freezeRows` | <span data-ttu-id="e6240-163">GeneralException</span><span class="sxs-lookup"><span data-stu-id="e6240-163">GeneralException</span></span> |
| `WorksheetFreezePanes.getLocationOrNullObject`| <span data-ttu-id="e6240-164">GeneralException</span><span class="sxs-lookup"><span data-stu-id="e6240-164">GeneralException</span></span> |
| `WorksheetFreezePanes.unfreeze` | <span data-ttu-id="e6240-165">GeneralException</span><span class="sxs-lookup"><span data-stu-id="e6240-165">GeneralException</span></span> |

> [!NOTE]
> <span data-ttu-id="e6240-166">Isso aplica-se apenas a várias pastas de trabalho do Excel abertas no Windows ou Mac.</span><span class="sxs-lookup"><span data-stu-id="e6240-166">This only applies to multiple Excel workbooks open on Windows or Mac.</span></span>

### <a name="coauthoring"></a><span data-ttu-id="e6240-167">Coautoria</span><span class="sxs-lookup"><span data-stu-id="e6240-167">Coauthoring</span></span>

<span data-ttu-id="e6240-168">Veja [coautoria em suplementos do Excel](../excel/co-authoring-in-excel-add-ins.md) para padrões a serem usados com eventos em um ambiente de coautoria.</span><span class="sxs-lookup"><span data-stu-id="e6240-168">See [Coauthoring in Excel add-ins](../excel/co-authoring-in-excel-add-ins.md) for patterns to use with events in a coauthoring environment.</span></span> <span data-ttu-id="e6240-169">O artigo também aborda possíveis conflitos de mesclagem ao usar determinadas APIs, como [`TableRowCollection.add`](/javascript/api/excel/excel.tablerowcollection#add-index--values-) .</span><span class="sxs-lookup"><span data-stu-id="e6240-169">The article also discusses potential merge conflicts when using certain APIs, such as [`TableRowCollection.add`](/javascript/api/excel/excel.tablerowcollection#add-index--values-).</span></span>

## <a name="see-also"></a><span data-ttu-id="e6240-170">Confira também</span><span class="sxs-lookup"><span data-stu-id="e6240-170">See also</span></span>

- [<span data-ttu-id="e6240-171">Limites de recurso e otimização de desempenho para Suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="e6240-171">Resource limits and performance optimization for Office Add-ins</span></span>](../concepts/resource-limits-and-performance-optimization.md)
- <span data-ttu-id="e6240-172">[OfficeDev/Office-js](https://github.com/OfficeDev/office-js/issues): o local para relatar e exibir problemas com a plataforma de suplementos do Office e APIs JavaScript.</span><span class="sxs-lookup"><span data-stu-id="e6240-172">[OfficeDev/office-js](https://github.com/OfficeDev/office-js/issues): The place to report and view issues with the Office Add-ins platform and JavaScript APIs.</span></span>
- <span data-ttu-id="e6240-173">[Estouro de pilha](https://stackoverflow.com/questions/tagged/office-js): o local para solicitar e exibir perguntas de programação sobre as APIs JavaScript do Office.</span><span class="sxs-lookup"><span data-stu-id="e6240-173">[Stack Overflow](https://stackoverflow.com/questions/tagged/office-js): The place to ask and view programming questions about the Office JavaScript APIs.</span></span> <span data-ttu-id="e6240-174">Certifique-se de aplicar a marca "Office-js" à sua pergunta ao postar no estouro de pilha.</span><span class="sxs-lookup"><span data-stu-id="e6240-174">Be sure to apply the "office-js" tag to your question when posting to Stack Overflow.</span></span>
- <span data-ttu-id="e6240-175">[UserVoice](https://officespdev.uservoice.com/): o local para sugerir novos recursos para a plataforma de suplementos do Office e APIs JavaScript do Office.</span><span class="sxs-lookup"><span data-stu-id="e6240-175">[UserVoice](https://officespdev.uservoice.com/): The place to suggest new features for the Office Add-ins platform and Office JavaScript APIs.</span></span>
