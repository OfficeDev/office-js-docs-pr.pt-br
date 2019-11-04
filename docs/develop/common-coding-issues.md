---
title: Problemas de codificação comuns e comportamentos de plataforma inesperados
description: Uma lista de problemas da plataforma de API JavaScript do Office frequentemente encontrada pelos desenvolvedores.
ms.date: 10/31/2019
localization_priority: Normal
ms.openlocfilehash: d39c379961833cdb924628becf2c2da3f7e271b9
ms.sourcegitcommit: 59d29d01bce7543ebebf86e5a86db00cf54ca14a
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 11/01/2019
ms.locfileid: "37924791"
---
# <a name="common-coding-issues-and-unexpected-platform-behaviors"></a><span data-ttu-id="db41c-103">Problemas de codificação comuns e comportamentos de plataforma inesperados</span><span class="sxs-lookup"><span data-stu-id="db41c-103">Common coding issues and unexpected platform behaviors</span></span>

<span data-ttu-id="db41c-104">Este artigo realça aspectos da API JavaScript do Office que podem resultar em comportamento inesperado ou exigir padrões de codificação específicos para obter o resultado desejado.</span><span class="sxs-lookup"><span data-stu-id="db41c-104">This article highlights aspects of the Office JavaScript API that may result in unexpected behavior or require specific coding patterns to achieve the desired outcome.</span></span> <span data-ttu-id="db41c-105">Se você encontrar um problema que pertença à lista, informe-nos usando o formulário de comentários na parte inferior do artigo.</span><span class="sxs-lookup"><span data-stu-id="db41c-105">If you encounter an issue that belongs in this list, please let us know by using the feedback form at the bottom of the article.</span></span>

## <a name="common-apis-and-outlook-apis-are-not-promise-based"></a><span data-ttu-id="db41c-106">APIs comuns e APIs do Outlook não são baseados em promessa</span><span class="sxs-lookup"><span data-stu-id="db41c-106">Common APIs and Outlook APIs are not promise-based</span></span>

<span data-ttu-id="db41c-107">As [APIs comuns](/javascript/api/office) (aquelas que não estão vinculadas a um host específico do Office) e [APIs do Outlook](/javascript/api/outlook) usam um modelo de programação baseado em retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="db41c-107">The [Common APIs](/javascript/api/office) (those that are not tied to a particular Office host) and [Outlook APIs](/javascript/api/outlook) use a callback-based programming model.</span></span> <span data-ttu-id="db41c-108">A interação com o documento subjacente do Office requer uma chamada de leitura ou gravação assíncrona que especifica um retorno de chamada a ser executado quando a operação for concluída.</span><span class="sxs-lookup"><span data-stu-id="db41c-108">Interacting with the underlying Office document requires an asynchronous read or write call that specifies a callback to be ran when the operation completes.</span></span> <span data-ttu-id="db41c-109">Para obter um exemplo desse padrão, consulte [Document. getFileAsync](/javascript/api/office/office.document#getfileasync-filetype--options--callback-).</span><span class="sxs-lookup"><span data-stu-id="db41c-109">For an example of this pattern, see [Document.getFileAsync](/javascript/api/office/office.document#getfileasync-filetype--options--callback-).</span></span>

<span data-ttu-id="db41c-110">Esses métodos comuns de API e API do Outlook não retornam [promessas](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise).</span><span class="sxs-lookup"><span data-stu-id="db41c-110">These Common API and Outlook API methods do not return [Promises](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise).</span></span> <span data-ttu-id="db41c-111">Portanto, você não pode usar [Await](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Operators/await) para pausar a execução até que a operação assíncrona seja concluída.</span><span class="sxs-lookup"><span data-stu-id="db41c-111">Therefore, you cannot use [await](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Operators/await) to pause the execution until the asynchronous operation completes.</span></span> <span data-ttu-id="db41c-112">Se você precisar `await` de comportamento, você pode encapsule a chamada do método em uma promessa criada explicitamente.</span><span class="sxs-lookup"><span data-stu-id="db41c-112">If you need `await` behavior, you can wrap the method call in an explicitly created Promise.</span></span>

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
> <span data-ttu-id="db41c-113">A documentação de referência contém a implementação com a promessa do [arquivo. getSliceAsync](/javascript/api/office/office.file#getsliceasync-sliceindex--callback-).</span><span class="sxs-lookup"><span data-stu-id="db41c-113">The reference documentation contains the Promise-wrapped implementation of [File.getSliceAsync](/javascript/api/office/office.file#getsliceasync-sliceindex--callback-).</span></span>

## <a name="some-properties-must-be-set-with-json-structs"></a><span data-ttu-id="db41c-114">Algumas propriedades devem ser definidas com as estruturas JSON</span><span class="sxs-lookup"><span data-stu-id="db41c-114">Some properties must be set with JSON structs</span></span>

> [!NOTE]
> <span data-ttu-id="db41c-115">Esta seção só se aplica às APIs específicas do host para Excel e Word.</span><span class="sxs-lookup"><span data-stu-id="db41c-115">This section only applies to the host-specific APIs for Excel and Word.</span></span>

<span data-ttu-id="db41c-116">Algumas propriedades devem ser definidas como estruturas JSON, em vez de definir suas subpropriedades individuais.</span><span class="sxs-lookup"><span data-stu-id="db41c-116">Some properties must be set as JSON structs, instead of setting their individual subproperties.</span></span> <span data-ttu-id="db41c-117">Um exemplo disso é encontrado no [PageLayout](/javascript/api/excel/excel.pagelayout).</span><span class="sxs-lookup"><span data-stu-id="db41c-117">One example of this is found in [PageLayout](/javascript/api/excel/excel.pagelayout).</span></span> <span data-ttu-id="db41c-118">A `zoom` propriedade deve ser definida com um único objeto [PageLayoutZoomOptions](/javascript/api/excel/excel.pagelayoutzoomoptions) , conforme mostrado aqui:</span><span class="sxs-lookup"><span data-stu-id="db41c-118">The `zoom` property must be set with a single [PageLayoutZoomOptions](/javascript/api/excel/excel.pagelayoutzoomoptions) object, as shown here:</span></span>

```js
// PageLayout.zoom must be set with JSON struct representing the PageLayoutZoomOptions object.
sheet.pageLayout.zoom = { scale: 200 };
```

<span data-ttu-id="db41c-119">No exemplo anterior, você ***não*** poderá atribuir `zoom` um valor diretamente: `sheet.pageLayout.zoom.scale = 200;`.</span><span class="sxs-lookup"><span data-stu-id="db41c-119">In the previous example, you would ***not*** be able to directly assign `zoom` a value: `sheet.pageLayout.zoom.scale = 200;`.</span></span> <span data-ttu-id="db41c-120">Essa instrução gera um erro porque `zoom` não está carregada.</span><span class="sxs-lookup"><span data-stu-id="db41c-120">That statement throws an error because `zoom` is not loaded.</span></span> <span data-ttu-id="db41c-121">Mesmo que `zoom` fosse carregado, o conjunto de escala não terá efeito.</span><span class="sxs-lookup"><span data-stu-id="db41c-121">Even if `zoom` were to be loaded, the set of scale will not take effect.</span></span> <span data-ttu-id="db41c-122">Todas as operações de contexto `zoom`acontecem em, atualizando o objeto de proxy no suplemento e substituindo os valores definidos localmente.</span><span class="sxs-lookup"><span data-stu-id="db41c-122">All context operations happen on `zoom`, refreshing the proxy object in the add-in and overwriting locally set values.</span></span>

<span data-ttu-id="db41c-123">Esse comportamento difere das [Propriedades de navegação](../excel/excel-add-ins-advanced-concepts.md#scalar-and-navigation-properties) , como [Range. Format](/javascript/api/excel/excel.range#format).</span><span class="sxs-lookup"><span data-stu-id="db41c-123">This behavior differs from [navigational properties](../excel/excel-add-ins-advanced-concepts.md#scalar-and-navigation-properties) like [Range.format](/javascript/api/excel/excel.range#format).</span></span> <span data-ttu-id="db41c-124">As propriedades `format` de podem ser definidas usando a navegação de objeto, conforme mostrado aqui:</span><span class="sxs-lookup"><span data-stu-id="db41c-124">Properties of `format` can be set using object navigation, as shown here:</span></span>

```js
// This will set the font size on the range during the next `content.sync()`.
range.format.font.size = 10;
```

<span data-ttu-id="db41c-125">Você pode identificar uma propriedade que deve ter suas subpropriedades definidas com uma estrutura JSON verificando seu modificador somente leitura.</span><span class="sxs-lookup"><span data-stu-id="db41c-125">You can identify a property that must have its subproperties set with a JSON struct by checking its read-only modifier.</span></span> <span data-ttu-id="db41c-126">Todas as propriedades somente leitura podem ter suas subpropriedades não somente leitura definidas diretamente.</span><span class="sxs-lookup"><span data-stu-id="db41c-126">All read-only properties can have their non-read-only subproperties directly set.</span></span> <span data-ttu-id="db41c-127">Propriedades graváveis como `PageLayout.zoom` devem ser definidas com uma estrutura JSON.</span><span class="sxs-lookup"><span data-stu-id="db41c-127">Writeable properties like `PageLayout.zoom` must be set with a JSON struct.</span></span> <span data-ttu-id="db41c-128">Em Resumo:</span><span class="sxs-lookup"><span data-stu-id="db41c-128">In summary:</span></span>

- <span data-ttu-id="db41c-129">Propriedade somente leitura: as subpropriedades podem ser definidas por meio de navegação.</span><span class="sxs-lookup"><span data-stu-id="db41c-129">Read-only property: Subproperties can be set through navigation.</span></span>
- <span data-ttu-id="db41c-130">Propriedade writable: as subpropriedades devem ser definidas com uma estrutura JSON (e não podem ser definidas por meio de navegação).</span><span class="sxs-lookup"><span data-stu-id="db41c-130">Writable property: Subproperties must be set with a JSON struct (and cannot be set through navigation).</span></span>

## <a name="excel-range-limits"></a><span data-ttu-id="db41c-131">Limites de intervalo do Excel</span><span class="sxs-lookup"><span data-stu-id="db41c-131">Excel Range limits</span></span>

<span data-ttu-id="db41c-132">Se você estiver criando um suplemento do Excel que usa intervalos, esteja ciente das seguintes limitações de tamanho:</span><span class="sxs-lookup"><span data-stu-id="db41c-132">If you're building an Excel add-in that uses ranges, be aware of the following size limitations:</span></span>

- <span data-ttu-id="db41c-133">O Excel na Web tem um limite de tamanho de conteúdo para solicitações e respostas de 5 MB.</span><span class="sxs-lookup"><span data-stu-id="db41c-133">Excel on the web has a payload size limit for requests and responses of 5MB.</span></span> <span data-ttu-id="db41c-134">`RichAPI.Error` será lançado se esse limite for excedido.</span><span class="sxs-lookup"><span data-stu-id="db41c-134">`RichAPI.Error` will be thrown if that limit is exceeded.</span></span>
- <span data-ttu-id="db41c-135">Um intervalo é limitado a 5 milhões células para operações de conjunto.</span><span class="sxs-lookup"><span data-stu-id="db41c-135">A range is limited to five million cells for set operations.</span></span>

<span data-ttu-id="db41c-136">Se você espera que a entrada do usuário exceda esses limites, verifique os dados e divida os intervalos em vários objetos.</span><span class="sxs-lookup"><span data-stu-id="db41c-136">If you expect user input to exceed these limits, be sure to check the data and split the ranges into multiple objects.</span></span> <span data-ttu-id="db41c-137">Você também precisará enviar várias `context.sync()` chamadas para evitar que as operações de intervalo menores fiquem novamente em lotes.</span><span class="sxs-lookup"><span data-stu-id="db41c-137">You'll also need to submit multiple `context.sync()` calls to avoid the smaller range operations getting batched together again.</span></span>

<span data-ttu-id="db41c-138">O suplemento pode ser capaz de usar o [RangeAreas](/javascript/api/excel/excel.rangeareas) para atualizar as células estrategicamente em um intervalo maior.</span><span class="sxs-lookup"><span data-stu-id="db41c-138">Your add-in might be able to use [RangeAreas](/javascript/api/excel/excel.rangeareas) to strategically update cells within a larger range.</span></span> <span data-ttu-id="db41c-139">Confira [trabalhar com vários intervalos simultaneamente em suplementos do Excel](../excel/excel-add-ins-multiple-ranges.md) para obter mais informações.</span><span class="sxs-lookup"><span data-stu-id="db41c-139">See [Work with multiple ranges simultaneously in Excel add-ins](../excel/excel-add-ins-multiple-ranges.md) for more information.</span></span>

## <a name="setting-read-only-properties"></a><span data-ttu-id="db41c-140">Configuração de propriedades somente leitura</span><span class="sxs-lookup"><span data-stu-id="db41c-140">Setting read-only properties</span></span>

<span data-ttu-id="db41c-141">As [definições do TypeScript](/referencing-the-javascript-api-for-office-library-from-its-cdn.md) para o Office js especificam quais propriedades de objeto são somente leitura.</span><span class="sxs-lookup"><span data-stu-id="db41c-141">The [TypeScript definitions](/referencing-the-javascript-api-for-office-library-from-its-cdn.md) for Office JS specify which object properties are read-only.</span></span> <span data-ttu-id="db41c-142">Se você tentar definir uma propriedade somente leitura, a operação de gravação falhará silenciosamente, sem nenhum erro.</span><span class="sxs-lookup"><span data-stu-id="db41c-142">If you attempt to set a read-only property, the write operation will fail silently, with no error thrown.</span></span> <span data-ttu-id="db41c-143">O exemplo a seguir tenta erroneamente definir a propriedade somente leitura [Chart.ID](/javascript/api/excel/excel.chart#id).</span><span class="sxs-lookup"><span data-stu-id="db41c-143">The following example erroneously attempts to set the read-only property [Chart.id](/javascript/api/excel/excel.chart#id).</span></span>

```js
// This will do nothing, since `id` is a read-only property.
myChart.id = "5";
```

## <a name="see-also"></a><span data-ttu-id="db41c-144">Confira também</span><span class="sxs-lookup"><span data-stu-id="db41c-144">See also</span></span>

- <span data-ttu-id="db41c-145">[OfficeDev/Office-js](https://github.com/OfficeDev/office-js/issues): o local para relatar e exibir problemas com a plataforma de suplementos do Office e APIs JavaScript.</span><span class="sxs-lookup"><span data-stu-id="db41c-145">[OfficeDev/office-js](https://github.com/OfficeDev/office-js/issues): The place to report and view issues with the Office Add-ins platform and JavaScript APIs.</span></span>
- <span data-ttu-id="db41c-146">[Estouro de pilha](https://stackoverflow.com/questions/tagged/office-js): o local para solicitar e exibir perguntas de programação sobre as APIs JavaScript do Office.</span><span class="sxs-lookup"><span data-stu-id="db41c-146">[Stack Overflow](https://stackoverflow.com/questions/tagged/office-js): The place to ask and view programming questions about the Office JavaScript APIs.</span></span> <span data-ttu-id="db41c-147">Certifique-se de aplicar a marca "Office-js" à sua pergunta ao postar no estouro de pilha.</span><span class="sxs-lookup"><span data-stu-id="db41c-147">Be sure to apply the "office-js" tag to your question when posting to Stack Overflow.</span></span>
- <span data-ttu-id="db41c-148">[UserVoice](https://officespdev.uservoice.com/): o local para sugerir novos recursos para a plataforma de suplementos do Office e APIs JavaScript do Office.</span><span class="sxs-lookup"><span data-stu-id="db41c-148">[UserVoice](https://officespdev.uservoice.com/): The place to suggest new features for the Office Add-ins platform and Office JavaScript APIs.</span></span>
