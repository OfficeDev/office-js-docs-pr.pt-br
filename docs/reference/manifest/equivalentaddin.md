---
title: Elemento EquivalentAddin no arquivo de manifesto
description: Especifica a compatibilidade COM versões anteriores para um suplemento COM equivalente ou XLL.
ms.date: 06/19/2019
localization_priority: Normal
ms.openlocfilehash: 425b926901b7325665eeede04263f74e4b854d50
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/17/2020
ms.locfileid: "42718283"
---
# <a name="equivalentaddin-element"></a><span data-ttu-id="e946d-103">Elemento EquivalentAddin</span><span class="sxs-lookup"><span data-stu-id="e946d-103">EquivalentAddin element</span></span>

<span data-ttu-id="e946d-104">Especifica a compatibilidade COM versões anteriores para um suplemento COM equivalente ou XLL.</span><span class="sxs-lookup"><span data-stu-id="e946d-104">Specifies backwards compatibility for an equivalent COM add-in or XLL.</span></span>

<span data-ttu-id="e946d-105">**Tipo de suplemento:** Painel de tarefas, função personalizada</span><span class="sxs-lookup"><span data-stu-id="e946d-105">**Add-in type:** Task pane, Custom function</span></span>

## <a name="syntax"></a><span data-ttu-id="e946d-106">Sintaxe</span><span class="sxs-lookup"><span data-stu-id="e946d-106">Syntax</span></span>

```XML
<EquivalentAddin>
   ...
</EquivalentAddin>
```

## <a name="contained-in"></a><span data-ttu-id="e946d-107">Contido em</span><span class="sxs-lookup"><span data-stu-id="e946d-107">Contained in</span></span>

[<span data-ttu-id="e946d-108">EquivalentAdd-ins</span><span class="sxs-lookup"><span data-stu-id="e946d-108">EquivalentAdd-ins</span></span>](equivalentaddins.md)

## <a name="must-contain"></a><span data-ttu-id="e946d-109">Deve conter</span><span class="sxs-lookup"><span data-stu-id="e946d-109">Must contain</span></span>

[<span data-ttu-id="e946d-110">Tipo</span><span class="sxs-lookup"><span data-stu-id="e946d-110">Type</span></span>](type.md)

## <a name="can-contain"></a><span data-ttu-id="e946d-111">Pode conter</span><span class="sxs-lookup"><span data-stu-id="e946d-111">Can contain</span></span>

<span data-ttu-id="e946d-112">[ProgId](progid.md)
[Nome de arquivo](filename.md) ProgID</span><span class="sxs-lookup"><span data-stu-id="e946d-112">[ProgId](progid.md)
[FileName](filename.md)</span></span>

## <a name="remarks"></a><span data-ttu-id="e946d-113">Comentários</span><span class="sxs-lookup"><span data-stu-id="e946d-113">Remarks</span></span>

<span data-ttu-id="e946d-114">Para especificar um suplemento de COM como o suplemento equivalente, forneça os `ProgId` elementos e. `Type`</span><span class="sxs-lookup"><span data-stu-id="e946d-114">To specify a COM add-in as the equivalent add-in, provide both the `ProgId` and `Type` elements.</span></span> <span data-ttu-id="e946d-115">Para especificar um XLL como o suplemento equivalente, forneça os `FileName` elementos e. `Type`</span><span class="sxs-lookup"><span data-stu-id="e946d-115">To specify an XLL as the equivalent add-in, provide both the `FileName` and `Type` elements.</span></span>

## <a name="see-also"></a><span data-ttu-id="e946d-116">Confira também</span><span class="sxs-lookup"><span data-stu-id="e946d-116">See also</span></span>

- [<span data-ttu-id="e946d-117">Torne as suas funções personalizadas compatíveis com as funções XLL definidas pelo usuário</span><span class="sxs-lookup"><span data-stu-id="e946d-117">Make your custom functions compatible with XLL user-defined functions</span></span>](../../excel/make-custom-functions-compatible-with-xll-udf.md)
- [<span data-ttu-id="e946d-118">Tornar seu suplemento do Excel compatível com um suplemento de COM existente</span><span class="sxs-lookup"><span data-stu-id="e946d-118">Make your Excel add-in compatible with an existing COM add-in</span></span>](../../develop/make-office-add-in-compatible-with-existing-com-add-in.md)