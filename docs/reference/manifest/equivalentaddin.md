---
title: Elemento EquivalentAddin no arquivo de manifesto
description: ''
ms.date: 06/19/2019
localization_priority: Normal
ms.openlocfilehash: 33cfb8b73e050fad7e392e0234962d346e903713
ms.sourcegitcommit: 4bf5159a3821f4277c07d89e88808c4c3a25ff81
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 06/18/2019
ms.locfileid: "35059920"
---
# <a name="equivalentaddin-element"></a><span data-ttu-id="cade4-102">Elemento EquivalentAddin</span><span class="sxs-lookup"><span data-stu-id="cade4-102">EquivalentAddin element</span></span>

<span data-ttu-id="cade4-103">Especifica a compatibilidade COM versões anteriores para um suplemento COM equivalente ou XLL.</span><span class="sxs-lookup"><span data-stu-id="cade4-103">Specifies backwards compatibility for an equivalent COM add-in or XLL.</span></span>

<span data-ttu-id="cade4-104">**Tipo de suplemento:** Painel de tarefas, função personalizada</span><span class="sxs-lookup"><span data-stu-id="cade4-104">**Add-in type:** Task pane, Custom function</span></span>

## <a name="syntax"></a><span data-ttu-id="cade4-105">Sintaxe</span><span class="sxs-lookup"><span data-stu-id="cade4-105">Syntax</span></span>

```XML
<EquivalentAddin>
   ...
</EquivalentAddin>
```

## <a name="contained-in"></a><span data-ttu-id="cade4-106">Contido em</span><span class="sxs-lookup"><span data-stu-id="cade4-106">Contained in</span></span>

[<span data-ttu-id="cade4-107">EquivalentAdd-ins</span><span class="sxs-lookup"><span data-stu-id="cade4-107">EquivalentAdd-ins</span></span>](equivalentaddins.md)

## <a name="must-contain"></a><span data-ttu-id="cade4-108">Deve conter</span><span class="sxs-lookup"><span data-stu-id="cade4-108">Must contain</span></span>

[<span data-ttu-id="cade4-109">Tipo</span><span class="sxs-lookup"><span data-stu-id="cade4-109">Type</span></span>](type.md)

## <a name="can-contain"></a><span data-ttu-id="cade4-110">Pode conter</span><span class="sxs-lookup"><span data-stu-id="cade4-110">Can contain</span></span>

<span data-ttu-id="cade4-111">[](progid.md)
[Nome de arquivo](filename.md) ProgID</span><span class="sxs-lookup"><span data-stu-id="cade4-111">[ProgId](progid.md)
[FileName](filename.md)</span></span>

## <a name="remarks"></a><span data-ttu-id="cade4-112">Comentários</span><span class="sxs-lookup"><span data-stu-id="cade4-112">Remarks</span></span>

<span data-ttu-id="cade4-113">Para especificar um suplemento de COM como o suplemento equivalente, forneça os `ProgId` elementos e. `Type`</span><span class="sxs-lookup"><span data-stu-id="cade4-113">To specify a COM add-in as the equivalent add-in, provide both the `ProgId` and `Type` elements.</span></span> <span data-ttu-id="cade4-114">Para especificar um XLL como o suplemento equivalente, forneça os `FileName` elementos e. `Type`</span><span class="sxs-lookup"><span data-stu-id="cade4-114">To specify an XLL as the equivalent add-in, provide both the `FileName` and `Type` elements.</span></span>

## <a name="see-also"></a><span data-ttu-id="cade4-115">Confira também</span><span class="sxs-lookup"><span data-stu-id="cade4-115">See also</span></span>

- [<span data-ttu-id="cade4-116">Torne as suas funções personalizadas compatíveis com as funções XLL definidas pelo usuário</span><span class="sxs-lookup"><span data-stu-id="cade4-116">Make your custom functions compatible with XLL user-defined functions</span></span>](../../excel/make-custom-functions-compatible-with-xll-udf.md)
- [<span data-ttu-id="cade4-117">Tornar seu suplemento do Excel compatível com um suplemento de COM existente</span><span class="sxs-lookup"><span data-stu-id="cade4-117">Make your Excel add-in compatible with an existing COM add-in</span></span>](../../develop/make-office-add-in-compatible-with-existing-com-add-in.md)