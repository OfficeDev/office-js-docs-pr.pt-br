---
title: O conjunto de requisitos somente online da API JavaScript do Excel
description: Detalhes sobre o conjunto de requisitos ExcelApiOnline.
ms.date: 09/15/2020
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 29f5826ba2adbf18b79033b83254b046210015fe
ms.sourcegitcommit: ed2a98b6fb5b432fa99c6cefa5ce52965dc25759
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 09/16/2020
ms.locfileid: "47819802"
---
# <a name="excel-javascript-api-online-only-requirement-set"></a><span data-ttu-id="7384b-103">O conjunto de requisitos somente online da API JavaScript do Excel</span><span class="sxs-lookup"><span data-stu-id="7384b-103">Excel JavaScript API online-only requirement set</span></span>

<span data-ttu-id="7384b-104">O `ExcelApiOnline` conjunto de requisitos é um conjunto de requisitos especiais que inclui recursos que estão disponíveis apenas para o Excel na Web.</span><span class="sxs-lookup"><span data-stu-id="7384b-104">The `ExcelApiOnline` requirement set is a special requirement set that includes features that are only available for Excel on the web.</span></span> <span data-ttu-id="7384b-105">As APIs neste conjunto de requisitos são consideradas APIs de produção (não sujeitas a alterações estruturais ou comportamentais não documentadas) para o Excel no aplicativo Web.</span><span class="sxs-lookup"><span data-stu-id="7384b-105">APIs in this requirement set are considered to be production APIs (not subject to undocumented behavioral or structural changes) for the Excel on the web application.</span></span> <span data-ttu-id="7384b-106">`ExcelApiOnline` são considerados como "Preview" APIs para outras plataformas (Windows, Mac, iOS) e podem não ser compatíveis com nenhuma dessas plataformas.</span><span class="sxs-lookup"><span data-stu-id="7384b-106">`ExcelApiOnline` are considered to be "preview" APIs for other platforms (Windows, Mac, iOS) and may not be supported by any of those platforms.</span></span>

<span data-ttu-id="7384b-107">Quando há suporte para APIs no `ExcelApiOnline` conjunto de requisitos em todas as plataformas, elas serão adicionadas ao próximo conjunto de requisitos liberados ( `ExcelApi 1.[NEXT]` ).</span><span class="sxs-lookup"><span data-stu-id="7384b-107">When APIs in the `ExcelApiOnline` requirement set are supported across all platforms, they will added to the next released requirement set (`ExcelApi 1.[NEXT]`).</span></span> <span data-ttu-id="7384b-108">Depois que o novo requisito for público, essas APIs serão removidas do `ExcelApiOnline` .</span><span class="sxs-lookup"><span data-stu-id="7384b-108">Once that new requirement is public, those APIs will be removed from `ExcelApiOnline`.</span></span> <span data-ttu-id="7384b-109">Pense nisso como um processo de promoção semelhante à de uma API que se move da versão prévia para o lançamento.</span><span class="sxs-lookup"><span data-stu-id="7384b-109">Think of this as a similar promotion process as an API moving from preview to release.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="7384b-110">`ExcelApiOnline` é o superconjunto do conjunto de requisitos mais recente.</span><span class="sxs-lookup"><span data-stu-id="7384b-110">`ExcelApiOnline` is superset of the latest numbered requirement set.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="7384b-111">`ExcelApiOnline 1.1` é a única versão das APIs somente online.</span><span class="sxs-lookup"><span data-stu-id="7384b-111">`ExcelApiOnline 1.1` is the only version of the online-only APIs.</span></span> <span data-ttu-id="7384b-112">Isso ocorre porque o Excel na Web sempre terá uma única versão disponível para os usuários que tenham a versão mais recente.</span><span class="sxs-lookup"><span data-stu-id="7384b-112">This is because Excel on the web will always have a single version available to users that is the latest version.</span></span>

## <a name="recommended-usage"></a><span data-ttu-id="7384b-113">Uso recomendado</span><span class="sxs-lookup"><span data-stu-id="7384b-113">Recommended usage</span></span>

<span data-ttu-id="7384b-114">Como as `ExcelApiOnline` APIs só têm suporte no Excel na Web, seu suplemento deve verificar se o conjunto de requisitos é suportado antes de chamar essas APIs.</span><span class="sxs-lookup"><span data-stu-id="7384b-114">Because `ExcelApiOnline` APIs are only supported by Excel on the web, your add-in should check if the requirement set is supported before calling these APIs.</span></span> <span data-ttu-id="7384b-115">Isso evita chamar uma API somente online em uma plataforma diferente.</span><span class="sxs-lookup"><span data-stu-id="7384b-115">This avoids calling an online-only API on a different platform.</span></span>

```js
if (Office.context.requirements.isSetSupported("ExcelApiOnline", "1.1")) {
   // Any API exclusive to the ExcelApiOnline requirement set.
}
```

<span data-ttu-id="7384b-116">Depois que a API estiver em um conjunto de requisitos de plataforma cruzada, você deverá remover ou editar a `isSetSupported` verificação.</span><span class="sxs-lookup"><span data-stu-id="7384b-116">Once the API is in a cross-platform requirement set, you should remove or edit the `isSetSupported` check.</span></span> <span data-ttu-id="7384b-117">Isso habilitará o recurso do seu suplemento em outras plataformas.</span><span class="sxs-lookup"><span data-stu-id="7384b-117">This will enable your add-in's feature on other platforms.</span></span> <span data-ttu-id="7384b-118">Certifique-se de testar o recurso nessas plataformas ao fazer essa alteração.</span><span class="sxs-lookup"><span data-stu-id="7384b-118">Be sure to test the feature on those platforms when making this change.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="7384b-119">O manifesto não pode ser especificado `ExcelApiOnline 1.1` como um requisito de ativação.</span><span class="sxs-lookup"><span data-stu-id="7384b-119">Your manifest cannot specify `ExcelApiOnline 1.1` as an activation requirement.</span></span> <span data-ttu-id="7384b-120">Não é um valor válido a ser usado no [elemento Set](../manifest/set.md).</span><span class="sxs-lookup"><span data-stu-id="7384b-120">It is not a valid value to use in the [Set element](../manifest/set.md).</span></span>

## <a name="api-list"></a><span data-ttu-id="7384b-121">Lista de APIs</span><span class="sxs-lookup"><span data-stu-id="7384b-121">API list</span></span>

<span data-ttu-id="7384b-122">No momento, não há nenhuma API no `ExcelApiOnline` conjunto de requisitos.</span><span class="sxs-lookup"><span data-stu-id="7384b-122">There are currently no APIs in the `ExcelApiOnline` requirement set.</span></span> <span data-ttu-id="7384b-123">Todas as APIs que anteriormente faziam parte deste conjunto graduaram para um conjunto de requisitos numerados e estão disponíveis em todas as plataformas.</span><span class="sxs-lookup"><span data-stu-id="7384b-123">All the APIs that were previously a part of this set have graduated to a numbered requirement set and are available across all platforms.</span></span>

## <a name="see-also"></a><span data-ttu-id="7384b-124">Confira também</span><span class="sxs-lookup"><span data-stu-id="7384b-124">See also</span></span>

- [<span data-ttu-id="7384b-125">Documentação deReferência da API JavaScript do Excel</span><span class="sxs-lookup"><span data-stu-id="7384b-125">Excel JavaScript API Reference Documentation</span></span>](/javascript/api/excel?view=excel-js-online&preserve-view=true)
- [<span data-ttu-id="7384b-126">APIs de visualização do JavaScript para Excel</span><span class="sxs-lookup"><span data-stu-id="7384b-126">Excel JavaScript preview APIs</span></span>](excel-preview-apis.md)
- [<span data-ttu-id="7384b-127">Conjuntos de requisitos da API JavaScript do Excel</span><span class="sxs-lookup"><span data-stu-id="7384b-127">Excel JavaScript API requirement sets</span></span>](excel-api-requirement-sets.md)
