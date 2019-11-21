---
title: O conjunto de requisitos somente online da API JavaScript do Excel
description: Detalhes sobre o conjunto de requisitos ExcelApiOnline
ms.date: 11/19/2019
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: e583c9832f04e17dc1c82d38d056fe2749888a77
ms.sourcegitcommit: e56bd8f1260c73daf33272a30dc5af242452594f
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 11/21/2019
ms.locfileid: "38757489"
---
# <a name="excel-javascript-api-online-only-requirement-set"></a><span data-ttu-id="fe534-103">O conjunto de requisitos somente online da API JavaScript do Excel</span><span class="sxs-lookup"><span data-stu-id="fe534-103">Excel JavaScript API online-only requirement set</span></span>

<span data-ttu-id="fe534-104">O `ExcelApiOnline` conjunto de requisitos é um conjunto de requisitos especiais que inclui recursos que estão disponíveis apenas para o Excel na Web.</span><span class="sxs-lookup"><span data-stu-id="fe534-104">The `ExcelApiOnline` requirement set is a special requirement set that includes features that are only available for Excel on the web.</span></span> <span data-ttu-id="fe534-105">As APIs neste conjunto de requisitos são consideradas APIs de produção (não sujeitas a alterações estruturais ou comportamentais não documentadas) para o Excel no host da Web.</span><span class="sxs-lookup"><span data-stu-id="fe534-105">APIs in this requirement set are considered to be production APIs (not subject to undocumented behavioral or structural changes) for the Excel on the web host.</span></span> <span data-ttu-id="fe534-106">`ExcelApiOnline`são considerados como "Preview" APIs para outras plataformas (Windows, Mac, iOS) e podem não ser compatíveis com nenhuma dessas plataformas.</span><span class="sxs-lookup"><span data-stu-id="fe534-106">`ExcelApiOnline` are considered to be "preview" APIs for other platforms (Windows, Mac, iOS) and may not be supported by any of those platforms.</span></span>

<span data-ttu-id="fe534-107">Quando há suporte para `ExcelApiOnline` APIs no conjunto de requisitos em todas as plataformas, elas serão adicionadas ao próximo conjunto de`ExcelApi 1.[NEXT]`requisitos liberados ().</span><span class="sxs-lookup"><span data-stu-id="fe534-107">When APIs in the `ExcelApiOnline` requirement set are supported across all platforms, they will added to the next released requirement set (`ExcelApi 1.[NEXT]`).</span></span> <span data-ttu-id="fe534-108">Depois que o novo requisito for público, essas APIs serão removidas do `ExcelApiOnline`.</span><span class="sxs-lookup"><span data-stu-id="fe534-108">Once that new requirement is public, those APIs will be removed from `ExcelApiOnline`.</span></span> <span data-ttu-id="fe534-109">Pense nisso como um processo de promoção semelhante à de uma API que se move da versão prévia para o lançamento.</span><span class="sxs-lookup"><span data-stu-id="fe534-109">Think of this as a similar promotion process as an API moving from preview to release.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="fe534-110">`ExcelApiOnline`é o superconjunto do conjunto de requisitos mais recente.</span><span class="sxs-lookup"><span data-stu-id="fe534-110">`ExcelApiOnline` is superset of the latest numbered requirement set.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="fe534-111">`ExcelApiOnline 1.1`é a única versão das APIs somente online.</span><span class="sxs-lookup"><span data-stu-id="fe534-111">`ExcelApiOnline 1.1` is the only version of the online-only APIs.</span></span> <span data-ttu-id="fe534-112">Isso ocorre porque o Excel na Web sempre terá uma única versão disponível para os usuários que tenham a versão mais recente.</span><span class="sxs-lookup"><span data-stu-id="fe534-112">This is because Excel on the web will always have a single version available to users that is the latest version.</span></span>

## <a name="recommended-usage"></a><span data-ttu-id="fe534-113">Uso recomendado</span><span class="sxs-lookup"><span data-stu-id="fe534-113">Recommended usage</span></span>

<span data-ttu-id="fe534-114">Como `ExcelApiOnline` as APIs só têm suporte no Excel na Web, seu suplemento deve verificar se o conjunto de requisitos é suportado antes de chamar essas APIs.</span><span class="sxs-lookup"><span data-stu-id="fe534-114">Because `ExcelApiOnline` APIs are only supported by Excel on the web, your add-in should check if the requirement set is supported before calling these APIs.</span></span> <span data-ttu-id="fe534-115">Isso evita chamar uma API somente online em uma plataforma diferente.</span><span class="sxs-lookup"><span data-stu-id="fe534-115">This avoids calling an online-only API on a different platform.</span></span>

```js
if (Office.context.requirements.isSetSupported("ExcelApiOnline", "1.1")) {
   // Any API exclusive to the ExcelApiOnline requirement set.
}
```

<span data-ttu-id="fe534-116">Depois que a API estiver em um conjunto de requisitos de plataforma cruzada, você deverá remover `isSetSupported` ou editar a verificação.</span><span class="sxs-lookup"><span data-stu-id="fe534-116">Once the API is in a cross-platform requirement set, you should remove or edit the `isSetSupported` check.</span></span> <span data-ttu-id="fe534-117">Isso habilitará o recurso do seu suplemento em outras plataformas.</span><span class="sxs-lookup"><span data-stu-id="fe534-117">This will enable your add-in's feature on other platforms.</span></span> <span data-ttu-id="fe534-118">Certifique-se de testar o recurso nessas plataformas ao fazer essa alteração.</span><span class="sxs-lookup"><span data-stu-id="fe534-118">Be sure to test the feature on those platforms when making this change.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="fe534-119">O manifesto não pode `ExcelApiOnline 1.1` ser especificado como um requisito de ativação.</span><span class="sxs-lookup"><span data-stu-id="fe534-119">Your manifest cannot specify `ExcelApiOnline 1.1` as an activation requirement.</span></span> <span data-ttu-id="fe534-120">Não é um valor válido a ser usado no [elemento Set](../manifest/set.md).</span><span class="sxs-lookup"><span data-stu-id="fe534-120">It is not a valid value to use in the [Set element](../manifest/set.md).</span></span>

## <a name="api-list"></a><span data-ttu-id="fe534-121">Lista de APIs</span><span class="sxs-lookup"><span data-stu-id="fe534-121">API list</span></span>

<span data-ttu-id="fe534-122">No momento, não há nenhuma API somente online.</span><span class="sxs-lookup"><span data-stu-id="fe534-122">There are currently no online-only APIs.</span></span> <span data-ttu-id="fe534-123">Confira novamente à medida que novos recursos são adicionados ao Excel na Web e suportados pelas APIs JavaScript do Office.</span><span class="sxs-lookup"><span data-stu-id="fe534-123">Check back as new features are added to Excel on the web and supported by the Office JavaScript APIs.</span></span>

## <a name="see-also"></a><span data-ttu-id="fe534-124">Confira também</span><span class="sxs-lookup"><span data-stu-id="fe534-124">See also</span></span>

- [<span data-ttu-id="fe534-125">Documentação deReferência da API JavaScript do Excel</span><span class="sxs-lookup"><span data-stu-id="fe534-125">Excel JavaScript API Reference Documentation</span></span>](/javascript/api/excel?view=excel-js-online)
- [<span data-ttu-id="fe534-126">APIs de visualização do JavaScript para Excel</span><span class="sxs-lookup"><span data-stu-id="fe534-126">Excel JavaScript preview APIs</span></span>](./excel-preview-apis.md)
- [<span data-ttu-id="fe534-127">Conjuntos de requisitos da API JavaScript do Excel</span><span class="sxs-lookup"><span data-stu-id="fe534-127">Excel JavaScript API requirement sets</span></span>](./excel-api-requirement-sets.md)