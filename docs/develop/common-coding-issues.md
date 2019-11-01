---
title: Problemas de codificação comuns e comportamentos de plataforma inesperados
description: Uma lista de problemas da plataforma de API JavaScript do Office frequentemente encontrada pelos desenvolvedores.
ms.date: 10/29/2019
localization_priority: Normal
ms.openlocfilehash: 8cea95e3214585ba8e0b77535916f9c564dde9df
ms.sourcegitcommit: e989096f3d19761bf8477c585cde20b3f8e0b90d
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 10/31/2019
ms.locfileid: "37902120"
---
# <a name="common-coding-issues-and-unexpected-platform-behaviors"></a><span data-ttu-id="9c961-103">Problemas de codificação comuns e comportamentos de plataforma inesperados</span><span class="sxs-lookup"><span data-stu-id="9c961-103">Common coding issues and unexpected platform behaviors</span></span>

<span data-ttu-id="9c961-104">Este artigo realça aspectos da API JavaScript do Office que podem resultar em comportamento inesperado ou exigir padrões de codificação específicos para obter o resultado desejado.</span><span class="sxs-lookup"><span data-stu-id="9c961-104">This article highlights aspects of the Office JavaScript API that may result in unexpected behavior or require specific coding patterns to achieve the desired outcome.</span></span> <span data-ttu-id="9c961-105">Se você encontrar um problema que pertença à lista, informe-nos usando o formulário de comentários na parte inferior do artigo.</span><span class="sxs-lookup"><span data-stu-id="9c961-105">If you encounter an issue that belongs in this list, please let us know by using the feedback form at the bottom of the article.</span></span>

## <a name="some-properties-must-be-set-with-json-structs"></a><span data-ttu-id="9c961-106">Algumas propriedades devem ser definidas com as estruturas JSON</span><span class="sxs-lookup"><span data-stu-id="9c961-106">Some properties must be set with JSON structs</span></span>

> [!NOTE]
> <span data-ttu-id="9c961-107">Esta seção só se aplica às APIs específicas do host para Excel e Word.</span><span class="sxs-lookup"><span data-stu-id="9c961-107">This section only applies to the host-specific APIs for Excel and Word.</span></span>

<span data-ttu-id="9c961-108">Algumas propriedades devem ser definidas como estruturas JSON, em vez de definir suas subpropriedades individuais.</span><span class="sxs-lookup"><span data-stu-id="9c961-108">Some properties must be set as JSON structs, instead of setting their individual subproperties.</span></span> <span data-ttu-id="9c961-109">Um exemplo disso é encontrado no [PageLayout](/javascript/api/excel/excel.pagelayout).</span><span class="sxs-lookup"><span data-stu-id="9c961-109">One example of this is found in [PageLayout](/javascript/api/excel/excel.pagelayout).</span></span> <span data-ttu-id="9c961-110">A `zoom` propriedade deve ser definida com um único objeto [PageLayoutZoomOptions](/javascript/api/excel/excel.pagelayoutzoomoptions) , conforme mostrado aqui:</span><span class="sxs-lookup"><span data-stu-id="9c961-110">The `zoom` property must be set with a single [PageLayoutZoomOptions](/javascript/api/excel/excel.pagelayoutzoomoptions) object, as shown here:</span></span>

```js
// PageLayout.zoom must be set with JSON struct representing the PageLayoutZoomOptions object.
sheet.pageLayout.zoom = { scale: 200 };
```

<span data-ttu-id="9c961-111">No exemplo anterior, você ***não*** poderá atribuir `zoom` um valor diretamente: `sheet.pageLayout.zoom.scale = 200;`.</span><span class="sxs-lookup"><span data-stu-id="9c961-111">In the previous example, you would ***not*** be able to directly assign `zoom` a value: `sheet.pageLayout.zoom.scale = 200;`.</span></span> <span data-ttu-id="9c961-112">Essa instrução gera um erro porque `zoom` não está carregada.</span><span class="sxs-lookup"><span data-stu-id="9c961-112">That statement throws an error because `zoom` is not loaded.</span></span> <span data-ttu-id="9c961-113">Mesmo que `zoom` fosse carregado, o conjunto de escala não terá efeito.</span><span class="sxs-lookup"><span data-stu-id="9c961-113">Even if `zoom` were to be loaded, the set of scale will not take effect.</span></span> <span data-ttu-id="9c961-114">Todas as operações de contexto `zoom`acontecem em, atualizando o objeto de proxy no suplemento e substituindo os valores definidos localmente.</span><span class="sxs-lookup"><span data-stu-id="9c961-114">All context operations happen on `zoom`, refreshing the proxy object in the add-in and overwriting locally set values.</span></span>

<span data-ttu-id="9c961-115">Esse comportamento difere das [Propriedades de navegação](../excel/excel-add-ins-advanced-concepts.md#scalar-and-navigation-properties) , como [Range. Format](/javascript/api/excel/excel.range#format).</span><span class="sxs-lookup"><span data-stu-id="9c961-115">This behavior differs from [navigational properties](../excel/excel-add-ins-advanced-concepts.md#scalar-and-navigation-properties) like [Range.format](/javascript/api/excel/excel.range#format).</span></span> <span data-ttu-id="9c961-116">As propriedades `format` de podem ser definidas usando a navegação de objeto, conforme mostrado aqui:</span><span class="sxs-lookup"><span data-stu-id="9c961-116">Properties of `format` can be set using object navigation, as shown here:</span></span>

```js
// This will set the font size on the range during the next `content.sync()`.
range.format.font.size = 10;
```

<span data-ttu-id="9c961-117">Você pode identificar uma propriedade que deve ter suas subpropriedades definidas com uma estrutura JSON verificando seu modificador somente leitura.</span><span class="sxs-lookup"><span data-stu-id="9c961-117">You can identify a property that must have its subproperties set with a JSON struct by checking its read-only modifier.</span></span> <span data-ttu-id="9c961-118">Todas as propriedades somente leitura podem ter suas subpropriedades não somente leitura definidas diretamente.</span><span class="sxs-lookup"><span data-stu-id="9c961-118">All read-only properties can have their non-read-only subproperties directly set.</span></span> <span data-ttu-id="9c961-119">Propriedades graváveis como `PageLayout.zoom` devem ser definidas com uma estrutura JSON.</span><span class="sxs-lookup"><span data-stu-id="9c961-119">Writeable properties like `PageLayout.zoom` must be set with a JSON struct.</span></span> <span data-ttu-id="9c961-120">Em Resumo:</span><span class="sxs-lookup"><span data-stu-id="9c961-120">In summary:</span></span>

- <span data-ttu-id="9c961-121">Propriedade somente leitura: as subpropriedades podem ser definidas por meio de navegação.</span><span class="sxs-lookup"><span data-stu-id="9c961-121">Read-only property: Subproperties can be set through navigation.</span></span>
- <span data-ttu-id="9c961-122">Propriedade writable: as subpropriedades devem ser definidas com uma estrutura JSON (e não podem ser definidas por meio de navegação).</span><span class="sxs-lookup"><span data-stu-id="9c961-122">Writable property: Subproperties must be set with a JSON struct (and cannot be set through navigation).</span></span>

## <a name="setting-read-only-properties"></a><span data-ttu-id="9c961-123">Configuração de propriedades somente leitura</span><span class="sxs-lookup"><span data-stu-id="9c961-123">Setting read-only properties</span></span>

<span data-ttu-id="9c961-124">As [definições do TypeScript](/referencing-the-javascript-api-for-office-library-from-its-cdn.md) para o Office js especificam quais propriedades de objeto são somente leitura.</span><span class="sxs-lookup"><span data-stu-id="9c961-124">The [TypeScript definitions](/referencing-the-javascript-api-for-office-library-from-its-cdn.md) for Office JS specify which object properties are read-only.</span></span> <span data-ttu-id="9c961-125">Se você tentar definir uma propriedade somente leitura, a operação de gravação falhará silenciosamente, sem nenhum erro.</span><span class="sxs-lookup"><span data-stu-id="9c961-125">If you attempt to set a read-only property, the write operation will fail silently, with no error thrown.</span></span> <span data-ttu-id="9c961-126">O exemplo a seguir tenta erroneamente definir a propriedade somente leitura [Chart.ID](/javascript/api/excel/excel.chart#id).</span><span class="sxs-lookup"><span data-stu-id="9c961-126">The following example erroneously attempts to set the read-only property [Chart.id](/javascript/api/excel/excel.chart#id).</span></span>

```js
// This will do nothing, since `id` is a read-only property.
myChart.id = "5";
```

## <a name="see-also"></a><span data-ttu-id="9c961-127">Confira também</span><span class="sxs-lookup"><span data-stu-id="9c961-127">See also</span></span>

- <span data-ttu-id="9c961-128">[OfficeDev/Office-js](https://github.com/OfficeDev/office-js/issues): o local para relatar e exibir problemas com a plataforma de suplementos do Office e APIs JavaScript.</span><span class="sxs-lookup"><span data-stu-id="9c961-128">[OfficeDev/office-js](https://github.com/OfficeDev/office-js/issues): The place to report and view issues with the Office Add-ins platform and JavaScript APIs.</span></span>
- <span data-ttu-id="9c961-129">[Estouro de pilha](https://stackoverflow.com/questions/tagged/office-js): o local para solicitar e exibir perguntas de programação sobre as APIs JavaScript do Office.</span><span class="sxs-lookup"><span data-stu-id="9c961-129">[Stack Overflow](https://stackoverflow.com/questions/tagged/office-js): The place to ask and view programming questions about the Office JavaScript APIs.</span></span> <span data-ttu-id="9c961-130">Certifique-se de aplicar a marca "Office-js" à sua pergunta ao postar no estouro de pilha.</span><span class="sxs-lookup"><span data-stu-id="9c961-130">Be sure to apply the "office-js" tag to your question when posting to Stack Overflow.</span></span>
- <span data-ttu-id="9c961-131">[UserVoice](https://officespdev.uservoice.com/): o local para sugerir novos recursos para a plataforma de suplementos do Office e APIs JavaScript do Office.</span><span class="sxs-lookup"><span data-stu-id="9c961-131">[UserVoice](https://officespdev.uservoice.com/): The place to suggest new features for the Office Add-ins platform and Office JavaScript APIs.</span></span>
