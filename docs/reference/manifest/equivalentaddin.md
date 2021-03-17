---
title: Elemento EquivalentAddin no arquivo de manifesto
description: Especifica a compatibilidade com vertida para um complemento COM ou XLL equivalente.
ms.date: 03/09/2021
localization_priority: Normal
ms.openlocfilehash: 412a3ce7bd12d886b7b88b5b84938e28295aba5d
ms.sourcegitcommit: c0c61fe84f3c5de88bd7eac29120056bb1224fc8
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/17/2021
ms.locfileid: "50836834"
---
# <a name="equivalentaddin-element"></a><span data-ttu-id="a439f-103">Elemento EquivalentAddin</span><span class="sxs-lookup"><span data-stu-id="a439f-103">EquivalentAddin element</span></span>

<span data-ttu-id="a439f-104">Especifica a compatibilidade com vertida para um complemento COM ou XLL equivalente.</span><span class="sxs-lookup"><span data-stu-id="a439f-104">Specifies backwards compatibility for an equivalent COM add-in or XLL.</span></span>

<span data-ttu-id="a439f-105">**Tipo de complemento:** Painel de tarefas, função Personalizada</span><span class="sxs-lookup"><span data-stu-id="a439f-105">**Add-in type:** Task pane, Custom function</span></span>

## <a name="syntax"></a><span data-ttu-id="a439f-106">Sintaxe</span><span class="sxs-lookup"><span data-stu-id="a439f-106">Syntax</span></span>

```XML
<EquivalentAddin>
   ...
</EquivalentAddin>
```

## <a name="contained-in"></a><span data-ttu-id="a439f-107">Contido em</span><span class="sxs-lookup"><span data-stu-id="a439f-107">Contained in</span></span>

[<span data-ttu-id="a439f-108">EquivalentAddins</span><span class="sxs-lookup"><span data-stu-id="a439f-108">EquivalentAddins</span></span>](equivalentaddins.md)

## <a name="must-contain"></a><span data-ttu-id="a439f-109">Deve conter</span><span class="sxs-lookup"><span data-stu-id="a439f-109">Must contain</span></span>

[<span data-ttu-id="a439f-110">Tipo</span><span class="sxs-lookup"><span data-stu-id="a439f-110">Type</span></span>](type.md)

## <a name="can-contain"></a><span data-ttu-id="a439f-111">Pode conter</span><span class="sxs-lookup"><span data-stu-id="a439f-111">Can contain</span></span>

<span data-ttu-id="a439f-112">[ProgId](progid.md) 
 [FileName](filename.md)</span><span class="sxs-lookup"><span data-stu-id="a439f-112">[ProgId](progid.md)
[FileName](filename.md)</span></span>

## <a name="remarks"></a><span data-ttu-id="a439f-113">Comentários</span><span class="sxs-lookup"><span data-stu-id="a439f-113">Remarks</span></span>

<span data-ttu-id="a439f-114">Para especificar um complemento COM como o complemento equivalente, forneça os `ProgId` elementos `Type` e.</span><span class="sxs-lookup"><span data-stu-id="a439f-114">To specify a COM add-in as the equivalent add-in, provide both the `ProgId` and `Type` elements.</span></span> <span data-ttu-id="a439f-115">Para especificar uma XLL como o complemento equivalente, forneça os `FileName` elementos `Type` e.</span><span class="sxs-lookup"><span data-stu-id="a439f-115">To specify an XLL as the equivalent add-in, provide both the `FileName` and `Type` elements.</span></span>

## <a name="see-also"></a><span data-ttu-id="a439f-116">Confira também</span><span class="sxs-lookup"><span data-stu-id="a439f-116">See also</span></span>

- [<span data-ttu-id="a439f-117">Torne as suas funções personalizadas compatíveis com as funções XLL definidas pelo usuário</span><span class="sxs-lookup"><span data-stu-id="a439f-117">Make your custom functions compatible with XLL user-defined functions</span></span>](../../excel/make-custom-functions-compatible-with-xll-udf.md)
- [<span data-ttu-id="a439f-118">Torne o seu suplemento do Office compatível com um suplemento COM existente</span><span class="sxs-lookup"><span data-stu-id="a439f-118">Make your Office Add-in compatible with an existing COM add-in</span></span>](../../develop/make-office-add-in-compatible-with-existing-com-add-in.md)