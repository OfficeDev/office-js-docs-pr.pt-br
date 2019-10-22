---
title: Elemento Type no arquivo de manifesto
description: ''
ms.date: 05/03/2019
localization_priority: Normal
ms.openlocfilehash: 1c053d65c5e3c6ce597c9912ec608e0b36bc623b
ms.sourcegitcommit: b3996b1444e520b44cf752e76eef50908386ca26
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 10/21/2019
ms.locfileid: "33628225"
---
# <a name="type-element"></a><span data-ttu-id="bc05c-102">Elemento Type</span><span class="sxs-lookup"><span data-stu-id="bc05c-102">Type element</span></span>

<span data-ttu-id="bc05c-103">Especifica se o suplemento equivalente é um suplemento COM ou um XLL.</span><span class="sxs-lookup"><span data-stu-id="bc05c-103">Specifies if the equivalent add-in is a COM addin or an XLL.</span></span>

<span data-ttu-id="bc05c-104">**Tipo de suplemento:** Painel de tarefas, função personalizada</span><span class="sxs-lookup"><span data-stu-id="bc05c-104">**Add-in type:** Task pane, Custom function</span></span>

## <a name="syntax"></a><span data-ttu-id="bc05c-105">Sintaxe</span><span class="sxs-lookup"><span data-stu-id="bc05c-105">Syntax</span></span>

```XML
    <Type> [COM | XLL] </Type>  
```

## <a name="contained-in"></a><span data-ttu-id="bc05c-106">Contido em</span><span class="sxs-lookup"><span data-stu-id="bc05c-106">Contained in</span></span>

[<span data-ttu-id="bc05c-107">EquivalentAdd-in</span><span class="sxs-lookup"><span data-stu-id="bc05c-107">EquivalentAdd-in</span></span>](equivalentaddin.md)

## <a name="add-in-type-values"></a><span data-ttu-id="bc05c-108">Valores de tipo de suplemento</span><span class="sxs-lookup"><span data-stu-id="bc05c-108">Add-in type values</span></span>

<span data-ttu-id="bc05c-109">Você deve especificar um dos seguintes valores para o `Type` elemento.</span><span class="sxs-lookup"><span data-stu-id="bc05c-109">You must specify one of the following values for the `Type` element.</span></span>

- <span data-ttu-id="bc05c-110">COM: especifica o suplemento equivalente é um suplemento de COM.</span><span class="sxs-lookup"><span data-stu-id="bc05c-110">COM: Specifies the equivalent add-in is a COM add-in.</span></span>
- <span data-ttu-id="bc05c-111">XLL: especifica o suplemento equivalente é um XLL do Excel.</span><span class="sxs-lookup"><span data-stu-id="bc05c-111">XLL: Specifies the equivalent add-in is an Excel XLL.</span></span>

## <a name="see-also"></a><span data-ttu-id="bc05c-112">Confira também</span><span class="sxs-lookup"><span data-stu-id="bc05c-112">See also</span></span>

- [<span data-ttu-id="bc05c-113">Torne as suas funções personalizadas compatíveis com as funções XLL definidas pelo usuário</span><span class="sxs-lookup"><span data-stu-id="bc05c-113">Make your custom functions compatible with XLL user-defined functions</span></span>](../../excel/make-custom-functions-compatible-with-xll-udf.md)
- [<span data-ttu-id="bc05c-114">Tornar seu suplemento do Excel compatível com um suplemento de COM existente</span><span class="sxs-lookup"><span data-stu-id="bc05c-114">Make your Excel add-in compatible with an existing COM add-in</span></span>](../../develop/make-office-add-in-compatible-with-existing-com-add-in.md)