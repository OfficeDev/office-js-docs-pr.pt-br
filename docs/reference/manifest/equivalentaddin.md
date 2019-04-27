---
title: Elemento EquivalentAddin no arquivo de manifesto
description: ''
ms.date: 04/22/2019
localization_priority: Normal
ms.openlocfilehash: 9cb1bb6d7a9cc3df3f4e39f8180b38d47d0a6882
ms.sourcegitcommit: 7462409209264dc7f8f89f3808a7a6249fcd739e
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 04/26/2019
ms.locfileid: "33356832"
---
# <a name="equivalentaddin-element"></a><span data-ttu-id="a6f5d-102">Elemento EquivalentAddin</span><span class="sxs-lookup"><span data-stu-id="a6f5d-102">EquivalentAddin element</span></span>

<span data-ttu-id="a6f5d-103">Especifica a compatibilidade COM versões anteriores para um suplemento COM equivalente ou XLL.</span><span class="sxs-lookup"><span data-stu-id="a6f5d-103">Specifies backwards compatibility for an equivalent COM add-in or XLL.</span></span>

<span data-ttu-id="a6f5d-104">**Tipo de suplemento:** Painel de tarefas, função personalizada</span><span class="sxs-lookup"><span data-stu-id="a6f5d-104">**Add-in type:** Task pane, Custom function</span></span>

## <a name="syntax"></a><span data-ttu-id="a6f5d-105">Sintaxe</span><span class="sxs-lookup"><span data-stu-id="a6f5d-105">Syntax</span></span>

```XML
<EquivalentAddin>
   ...
</EquivalentAddin>
```

## <a name="contained-in"></a><span data-ttu-id="a6f5d-106">Contido em</span><span class="sxs-lookup"><span data-stu-id="a6f5d-106">Contained in</span></span>

[<span data-ttu-id="a6f5d-107">EquivalentAdd-ins</span><span class="sxs-lookup"><span data-stu-id="a6f5d-107">EquivalentAdd-ins</span></span>](equivalentaddins.md)

## <a name="must-contain"></a><span data-ttu-id="a6f5d-108">Deve conter</span><span class="sxs-lookup"><span data-stu-id="a6f5d-108">Must contain</span></span>

[<span data-ttu-id="a6f5d-109">Type</span><span class="sxs-lookup"><span data-stu-id="a6f5d-109">Type</span></span>](type.md)

## <a name="can-contain"></a><span data-ttu-id="a6f5d-110">Pode conter</span><span class="sxs-lookup"><span data-stu-id="a6f5d-110">Can contain</span></span>

<span data-ttu-id="a6f5d-111">[](progid.md)
[Nome de arquivo](filename.md) ProgID</span><span class="sxs-lookup"><span data-stu-id="a6f5d-111">[ProgID](progid.md)
[FileName](filename.md)</span></span>

## <a name="remarks"></a><span data-ttu-id="a6f5d-112">Comentários</span><span class="sxs-lookup"><span data-stu-id="a6f5d-112">Remarks</span></span>

<span data-ttu-id="a6f5d-113">Para especificar um suplemento de COM como o suplemento equivalente, forneça os `ProgID` elementos e. `Type`</span><span class="sxs-lookup"><span data-stu-id="a6f5d-113">To specify a COM add-in as the equivalent add-in, provide both the `ProgID` and `Type` elements.</span></span> <span data-ttu-id="a6f5d-114">Para especificar um XLL como o suplemento equivalente, forneça os `FileName` elementos e. `Type`</span><span class="sxs-lookup"><span data-stu-id="a6f5d-114">To specify an XLL as the equivalent add-in, provide both the `FileName` and `Type` elements.</span></span>

## <a name="see-also"></a><span data-ttu-id="a6f5d-115">Confira também</span><span class="sxs-lookup"><span data-stu-id="a6f5d-115">See also</span></span>

- [<span data-ttu-id="a6f5d-116">Tornar suas funções personalizadas compatíveis com as funções definidas pelo usuário XLL</span><span class="sxs-lookup"><span data-stu-id="a6f5d-116">Make your custom functions compatible with XLL user-defined functions</span></span>](../../excel/make-custom-functions-compatible-with-xll-udf.md)
- [<span data-ttu-id="a6f5d-117">Tornar o suplemento do Office compatível com um suplemento de COM existente</span><span class="sxs-lookup"><span data-stu-id="a6f5d-117">Make your Office Add-in compatible with an existing COM add-in</span></span>](../../develop/make-office-add-in-compatible-with-existing-com-add-in.md)