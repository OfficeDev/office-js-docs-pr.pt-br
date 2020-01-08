---
title: Problemas de codificação comuns e comportamentos de plataforma inesperados
description: Uma lista de problemas da plataforma de API JavaScript do Office frequentemente encontrada pelos desenvolvedores.
ms.date: 01/02/2020
localization_priority: Normal
ms.openlocfilehash: fa33451550ab02f76a8b41ebf682e6a73d2a3a96
ms.sourcegitcommit: abe8188684b55710261c69e206de83d3a6bd2ed3
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 01/08/2020
ms.locfileid: "40969490"
---
# <a name="common-coding-issues-and-unexpected-platform-behaviors"></a><span data-ttu-id="69ef0-103">Problemas de codificação comuns e comportamentos de plataforma inesperados</span><span class="sxs-lookup"><span data-stu-id="69ef0-103">Common coding issues and unexpected platform behaviors</span></span>

<span data-ttu-id="69ef0-104">Este artigo realça aspectos da API JavaScript do Office que podem resultar em comportamento inesperado ou exigir padrões de codificação específicos para obter o resultado desejado.</span><span class="sxs-lookup"><span data-stu-id="69ef0-104">This article highlights aspects of the Office JavaScript API that may result in unexpected behavior or require specific coding patterns to achieve the desired outcome.</span></span> <span data-ttu-id="69ef0-105">Se você encontrar um problema que pertença à lista, informe-nos usando o formulário de comentários na parte inferior do artigo.</span><span class="sxs-lookup"><span data-stu-id="69ef0-105">If you encounter an issue that belongs in this list, please let us know by using the feedback form at the bottom of the article.</span></span>

## <a name="common-apis-and-outlook-apis-are-not-promise-based"></a><span data-ttu-id="69ef0-106">APIs comuns e APIs do Outlook não são baseados em promessa</span><span class="sxs-lookup"><span data-stu-id="69ef0-106">Common APIs and Outlook APIs are not promise-based</span></span>

<span data-ttu-id="69ef0-107">As [APIs comuns](/javascript/api/office) (aquelas que não estão vinculadas a um host específico do Office) e [APIs do Outlook](/javascript/api/outlook) usam um modelo de programação baseado em retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="69ef0-107">The [Common APIs](/javascript/api/office) (those that are not tied to a particular Office host) and [Outlook APIs](/javascript/api/outlook) use a callback-based programming model.</span></span> <span data-ttu-id="69ef0-108">A interação com o documento subjacente do Office requer uma chamada de leitura ou gravação assíncrona que especifica um retorno de chamada a ser executado quando a operação for concluída.</span><span class="sxs-lookup"><span data-stu-id="69ef0-108">Interacting with the underlying Office document requires an asynchronous read or write call that specifies a callback to be ran when the operation completes.</span></span> <span data-ttu-id="69ef0-109">Para obter um exemplo desse padrão, consulte [Document. getFileAsync](/javascript/api/office/office.document#getfileasync-filetype--options--callback-).</span><span class="sxs-lookup"><span data-stu-id="69ef0-109">For an example of this pattern, see [Document.getFileAsync](/javascript/api/office/office.document#getfileasync-filetype--options--callback-).</span></span>

<span data-ttu-id="69ef0-110">Esses métodos comuns de API e API do Outlook não retornam [promessas](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise).</span><span class="sxs-lookup"><span data-stu-id="69ef0-110">These Common API and Outlook API methods do not return [Promises](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise).</span></span> <span data-ttu-id="69ef0-111">Portanto, você não pode usar [Await](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Operators/await) para pausar a execução até que a operação assíncrona seja concluída.</span><span class="sxs-lookup"><span data-stu-id="69ef0-111">Therefore, you cannot use [await](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Operators/await) to pause the execution until the asynchronous operation completes.</span></span> <span data-ttu-id="69ef0-112">Se você precisar `await` de comportamento, você pode encapsule a chamada do método em uma promessa criada explicitamente.</span><span class="sxs-lookup"><span data-stu-id="69ef0-112">If you need `await` behavior, you can wrap the method call in an explicitly created Promise.</span></span>

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
> <span data-ttu-id="69ef0-113">A documentação de referência contém a implementação com a promessa do [arquivo. getSliceAsync](/javascript/api/office/office.file#getsliceasync-sliceindex--callback-).</span><span class="sxs-lookup"><span data-stu-id="69ef0-113">The reference documentation contains the Promise-wrapped implementation of [File.getSliceAsync](/javascript/api/office/office.file#getsliceasync-sliceindex--callback-).</span></span>

## <a name="some-properties-cannot-be-set-directly"></a><span data-ttu-id="69ef0-114">Algumas propriedades não podem ser definidas diretamente</span><span class="sxs-lookup"><span data-stu-id="69ef0-114">Some properties cannot be set directly</span></span>

> [!NOTE]
> <span data-ttu-id="69ef0-115">Esta seção só se aplica às APIs específicas do host para Excel e Word.</span><span class="sxs-lookup"><span data-stu-id="69ef0-115">This section only applies to the host-specific APIs for Excel and Word.</span></span>

<span data-ttu-id="69ef0-116">Algumas propriedades não podem ser definidas, apesar de serem graváveis.</span><span class="sxs-lookup"><span data-stu-id="69ef0-116">Some properties cannot be set, despite being writable.</span></span> <span data-ttu-id="69ef0-117">Essas propriedades fazem parte de uma propriedade pai que deve ser definida como um único objeto.</span><span class="sxs-lookup"><span data-stu-id="69ef0-117">These properties are part of a parent property that must be set as a single object.</span></span> <span data-ttu-id="69ef0-118">Isso ocorre porque a propriedade Parent depende das subpropriedades que têm relações lógicas específicas.</span><span class="sxs-lookup"><span data-stu-id="69ef0-118">This is because that parent property relies on the subproperties having specific, logical relationships.</span></span> <span data-ttu-id="69ef0-119">Essas propriedades pai devem ser definidas usando a notação literal de objeto para definir o objeto inteiro, em vez de definir as subpropriedades individuais desse objeto.</span><span class="sxs-lookup"><span data-stu-id="69ef0-119">These parent properties must be set using object literal notation to set the entire object, instead of setting that object's individual subproperties.</span></span> <span data-ttu-id="69ef0-120">Um exemplo disso é encontrado no [PageLayout](/javascript/api/excel/excel.pagelayout).</span><span class="sxs-lookup"><span data-stu-id="69ef0-120">One example of this is found in [PageLayout](/javascript/api/excel/excel.pagelayout).</span></span> <span data-ttu-id="69ef0-121">A `zoom` propriedade deve ser definida com um único objeto [PageLayoutZoomOptions](/javascript/api/excel/excel.pagelayoutzoomoptions) , conforme mostrado aqui:</span><span class="sxs-lookup"><span data-stu-id="69ef0-121">The `zoom` property must be set with a single [PageLayoutZoomOptions](/javascript/api/excel/excel.pagelayoutzoomoptions) object, as shown here:</span></span>

```js
// PageLayout.zoom.scale must be set by assigning PageLayout.zoom to a PageLayoutZoomOptions object.
sheet.pageLayout.zoom = { scale: 200 };
```

<span data-ttu-id="69ef0-122">No exemplo anterior, você ***não*** poderá atribuir `zoom` um valor diretamente: `sheet.pageLayout.zoom.scale = 200;`.</span><span class="sxs-lookup"><span data-stu-id="69ef0-122">In the previous example, you would ***not*** be able to directly assign `zoom` a value: `sheet.pageLayout.zoom.scale = 200;`.</span></span> <span data-ttu-id="69ef0-123">Essa instrução gera um erro porque `zoom` não está carregada.</span><span class="sxs-lookup"><span data-stu-id="69ef0-123">That statement throws an error because `zoom` is not loaded.</span></span> <span data-ttu-id="69ef0-124">Mesmo que `zoom` fosse carregado, o conjunto de escala não terá efeito.</span><span class="sxs-lookup"><span data-stu-id="69ef0-124">Even if `zoom` were to be loaded, the set of scale will not take effect.</span></span> <span data-ttu-id="69ef0-125">Todas as operações de contexto `zoom`acontecem em, atualizando o objeto de proxy no suplemento e substituindo os valores definidos localmente.</span><span class="sxs-lookup"><span data-stu-id="69ef0-125">All context operations happen on `zoom`, refreshing the proxy object in the add-in and overwriting locally set values.</span></span>

<span data-ttu-id="69ef0-126">Esse comportamento difere das [Propriedades de navegação](../excel/excel-add-ins-advanced-concepts.md#scalar-and-navigation-properties) , como [Range. Format](/javascript/api/excel/excel.range#format).</span><span class="sxs-lookup"><span data-stu-id="69ef0-126">This behavior differs from [navigational properties](../excel/excel-add-ins-advanced-concepts.md#scalar-and-navigation-properties) like [Range.format](/javascript/api/excel/excel.range#format).</span></span> <span data-ttu-id="69ef0-127">As propriedades `format` de podem ser definidas usando a navegação de objeto, conforme mostrado aqui:</span><span class="sxs-lookup"><span data-stu-id="69ef0-127">Properties of `format` can be set using object navigation, as shown here:</span></span>

```js
// This will set the font size on the range during the next `content.sync()`.
range.format.font.size = 10;
```

<span data-ttu-id="69ef0-128">Você pode identificar uma propriedade que não pode ter suas subpropriedades definidas diretamente verificando seu modificador somente leitura.</span><span class="sxs-lookup"><span data-stu-id="69ef0-128">You can identify a property that cannot have its subproperties directly set by checking its read-only modifier.</span></span> <span data-ttu-id="69ef0-129">Todas as propriedades somente leitura podem ter suas subpropriedades não somente leitura definidas diretamente.</span><span class="sxs-lookup"><span data-stu-id="69ef0-129">All read-only properties can have their non-read-only subproperties directly set.</span></span> <span data-ttu-id="69ef0-130">Propriedades graváveis como `PageLayout.zoom` devem ser definidas com um objeto nesse nível.</span><span class="sxs-lookup"><span data-stu-id="69ef0-130">Writeable properties like `PageLayout.zoom` must be set with an object at that level.</span></span> <span data-ttu-id="69ef0-131">Em Resumo:</span><span class="sxs-lookup"><span data-stu-id="69ef0-131">In summary:</span></span>

- <span data-ttu-id="69ef0-132">Propriedade somente leitura: as subpropriedades podem ser definidas por meio de navegação.</span><span class="sxs-lookup"><span data-stu-id="69ef0-132">Read-only property: Subproperties can be set through navigation.</span></span>
- <span data-ttu-id="69ef0-133">Propriedade writable: as subpropriedades não podem ser definidas por meio de navegação (devem ser definidas como parte da atribuição de objeto pai inicial).</span><span class="sxs-lookup"><span data-stu-id="69ef0-133">Writable property: Subproperties cannot be set through navigation (must be set as part of the initial parent object assignment).</span></span>

## <a name="excel-data-transfer-limits"></a><span data-ttu-id="69ef0-134">Limites de transferência de dados do Excel</span><span class="sxs-lookup"><span data-stu-id="69ef0-134">Excel data transfer limits</span></span>

<span data-ttu-id="69ef0-135">Se você estiver criando um suplemento do Excel, esteja ciente das seguintes limitações de tamanho ao interagir com a pasta de trabalho:</span><span class="sxs-lookup"><span data-stu-id="69ef0-135">If you're building an Excel add-in, be aware of the following size limitations when interacting with the workbook:</span></span>

- <span data-ttu-id="69ef0-136">O Excel na Web tem um limite de tamanho de conteúdo para solicitações e respostas de 5 MB.</span><span class="sxs-lookup"><span data-stu-id="69ef0-136">Excel on the web has a payload size limit for requests and responses of 5MB.</span></span> <span data-ttu-id="69ef0-137">`RichAPI.Error` será lançado se esse limite for excedido.</span><span class="sxs-lookup"><span data-stu-id="69ef0-137">`RichAPI.Error` will be thrown if that limit is exceeded.</span></span>
- <span data-ttu-id="69ef0-138">Um intervalo está limitado a 5 milhões células para operações Get.</span><span class="sxs-lookup"><span data-stu-id="69ef0-138">A range is limited to five million cells for get operations.</span></span>

<span data-ttu-id="69ef0-139">Se você espera que a entrada do usuário exceda esses limites, verifique os dados antes de `context.sync()`chamar.</span><span class="sxs-lookup"><span data-stu-id="69ef0-139">If you expect user input to exceed these limits, be sure to check the data before calling `context.sync()`.</span></span> <span data-ttu-id="69ef0-140">Divida a operação em partes menores, conforme necessário.</span><span class="sxs-lookup"><span data-stu-id="69ef0-140">Split the operation into smaller pieces as needed.</span></span> <span data-ttu-id="69ef0-141">Certifique-se de `context.sync()` chamar para cada suboperação para evitar que as operações sejam encaixadas novamente.</span><span class="sxs-lookup"><span data-stu-id="69ef0-141">Be sure to call `context.sync()` for each sub-operation to avoid those operations getting batched together again.</span></span>

<span data-ttu-id="69ef0-142">Essas limitações são normalmente excedidos por intervalos grandes.</span><span class="sxs-lookup"><span data-stu-id="69ef0-142">These limitations are typically exceeded by large ranges.</span></span> <span data-ttu-id="69ef0-143">O suplemento pode ser capaz de usar o [RangeAreas](/javascript/api/excel/excel.rangeareas) para atualizar as células estrategicamente em um intervalo maior.</span><span class="sxs-lookup"><span data-stu-id="69ef0-143">Your add-in might be able to use [RangeAreas](/javascript/api/excel/excel.rangeareas) to strategically update cells within a larger range.</span></span> <span data-ttu-id="69ef0-144">Confira [trabalhar com vários intervalos simultaneamente em suplementos do Excel](../excel/excel-add-ins-multiple-ranges.md) para obter mais informações.</span><span class="sxs-lookup"><span data-stu-id="69ef0-144">See [Work with multiple ranges simultaneously in Excel add-ins](../excel/excel-add-ins-multiple-ranges.md) for more information.</span></span>

## <a name="setting-read-only-properties"></a><span data-ttu-id="69ef0-145">Configuração de propriedades somente leitura</span><span class="sxs-lookup"><span data-stu-id="69ef0-145">Setting read-only properties</span></span>

<span data-ttu-id="69ef0-146">As [definições do TypeScript](referencing-the-javascript-api-for-office-library-from-its-cdn.md) para o Office js especificam quais propriedades de objeto são somente leitura.</span><span class="sxs-lookup"><span data-stu-id="69ef0-146">The [TypeScript definitions](referencing-the-javascript-api-for-office-library-from-its-cdn.md) for Office JS specify which object properties are read-only.</span></span> <span data-ttu-id="69ef0-147">Se você tentar definir uma propriedade somente leitura, a operação de gravação falhará silenciosamente, sem nenhum erro.</span><span class="sxs-lookup"><span data-stu-id="69ef0-147">If you attempt to set a read-only property, the write operation will fail silently, with no error thrown.</span></span> <span data-ttu-id="69ef0-148">O exemplo a seguir tenta erroneamente definir a propriedade somente leitura [Chart.ID](/javascript/api/excel/excel.chart#id).</span><span class="sxs-lookup"><span data-stu-id="69ef0-148">The following example erroneously attempts to set the read-only property [Chart.id](/javascript/api/excel/excel.chart#id).</span></span>

```js
// This will do nothing, since `id` is a read-only property.
myChart.id = "5";
```

## <a name="removing-event-handlers"></a><span data-ttu-id="69ef0-149">Remover manipuladores de eventos</span><span class="sxs-lookup"><span data-stu-id="69ef0-149">Removing event handlers</span></span>

<span data-ttu-id="69ef0-150">Manipuladores de eventos devem ser removidos usando o `RequestContext` mesmo em que foram adicionados.</span><span class="sxs-lookup"><span data-stu-id="69ef0-150">Event handlers must be removed using the same `RequestContext` in which they were added.</span></span> <span data-ttu-id="69ef0-151">Se você precisar que seu suplemento remova um manipulador de eventos durante a execução, será necessário armazenar o objeto Context usado para adicionar o manipulador.</span><span class="sxs-lookup"><span data-stu-id="69ef0-151">If you need your add-in to remove an event handler while running, you'll need to store the context object used to add the handler.</span></span>

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

## <a name="see-also"></a><span data-ttu-id="69ef0-152">Confira também</span><span class="sxs-lookup"><span data-stu-id="69ef0-152">See also</span></span>

- <span data-ttu-id="69ef0-153">[OfficeDev/Office-js](https://github.com/OfficeDev/office-js/issues): o local para relatar e exibir problemas com a plataforma de suplementos do Office e APIs JavaScript.</span><span class="sxs-lookup"><span data-stu-id="69ef0-153">[OfficeDev/office-js](https://github.com/OfficeDev/office-js/issues): The place to report and view issues with the Office Add-ins platform and JavaScript APIs.</span></span>
- <span data-ttu-id="69ef0-154">[Estouro de pilha](https://stackoverflow.com/questions/tagged/office-js): o local para solicitar e exibir perguntas de programação sobre as APIs JavaScript do Office.</span><span class="sxs-lookup"><span data-stu-id="69ef0-154">[Stack Overflow](https://stackoverflow.com/questions/tagged/office-js): The place to ask and view programming questions about the Office JavaScript APIs.</span></span> <span data-ttu-id="69ef0-155">Certifique-se de aplicar a marca "Office-js" à sua pergunta ao postar no estouro de pilha.</span><span class="sxs-lookup"><span data-stu-id="69ef0-155">Be sure to apply the "office-js" tag to your question when posting to Stack Overflow.</span></span>
- <span data-ttu-id="69ef0-156">[UserVoice](https://officespdev.uservoice.com/): o local para sugerir novos recursos para a plataforma de suplementos do Office e APIs JavaScript do Office.</span><span class="sxs-lookup"><span data-stu-id="69ef0-156">[UserVoice](https://officespdev.uservoice.com/): The place to suggest new features for the Office Add-ins platform and Office JavaScript APIs.</span></span>
