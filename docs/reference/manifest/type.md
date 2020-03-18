---
title: Elemento Type no arquivo de manifesto
description: O elemento Type Especifica se o suplemento equivalente é um suplemento de COM ou um XLL.
ms.date: 03/16/2020
localization_priority: Normal
ms.openlocfilehash: 9eeab172ed4ebf06fc93e42f56f8d33f5e7a92db
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/17/2020
ms.locfileid: "42720313"
---
# <a name="type-element"></a><span data-ttu-id="64b27-103">Elemento Type</span><span class="sxs-lookup"><span data-stu-id="64b27-103">Type element</span></span>

<span data-ttu-id="64b27-104">Especifica se o suplemento equivalente é um suplemento de COM ou um XLL.</span><span class="sxs-lookup"><span data-stu-id="64b27-104">Specifies if the equivalent add-in is a COM add-in or an XLL.</span></span>

<span data-ttu-id="64b27-105">**Tipo de suplemento:** Painel de tarefas, função personalizada</span><span class="sxs-lookup"><span data-stu-id="64b27-105">**Add-in type:** Task pane, Custom function</span></span>

## <a name="syntax"></a><span data-ttu-id="64b27-106">Sintaxe</span><span class="sxs-lookup"><span data-stu-id="64b27-106">Syntax</span></span>

```XML
    <Type> [COM | XLL] </Type>  
```

## <a name="contained-in"></a><span data-ttu-id="64b27-107">Contido em</span><span class="sxs-lookup"><span data-stu-id="64b27-107">Contained in</span></span>

[<span data-ttu-id="64b27-108">EquivalentAdd-in</span><span class="sxs-lookup"><span data-stu-id="64b27-108">EquivalentAdd-in</span></span>](equivalentaddin.md)

## <a name="add-in-type-values"></a><span data-ttu-id="64b27-109">Valores de tipo de suplemento</span><span class="sxs-lookup"><span data-stu-id="64b27-109">Add-in type values</span></span>

<span data-ttu-id="64b27-110">Você deve especificar um dos seguintes valores para o `Type` elemento.</span><span class="sxs-lookup"><span data-stu-id="64b27-110">You must specify one of the following values for the `Type` element.</span></span>

- <span data-ttu-id="64b27-111">COM: especifica o suplemento equivalente é um suplemento de COM.</span><span class="sxs-lookup"><span data-stu-id="64b27-111">COM: Specifies the equivalent add-in is a COM add-in.</span></span>
- <span data-ttu-id="64b27-112">XLL: especifica o suplemento equivalente é um XLL do Excel.</span><span class="sxs-lookup"><span data-stu-id="64b27-112">XLL: Specifies the equivalent add-in is an Excel XLL.</span></span>

## <a name="see-also"></a><span data-ttu-id="64b27-113">Confira também</span><span class="sxs-lookup"><span data-stu-id="64b27-113">See also</span></span>

- [<span data-ttu-id="64b27-114">Torne as suas funções personalizadas compatíveis com as funções XLL definidas pelo usuário</span><span class="sxs-lookup"><span data-stu-id="64b27-114">Make your custom functions compatible with XLL user-defined functions</span></span>](../../excel/make-custom-functions-compatible-with-xll-udf.md)
- [<span data-ttu-id="64b27-115">Tornar seu suplemento do Excel compatível com um suplemento de COM existente</span><span class="sxs-lookup"><span data-stu-id="64b27-115">Make your Excel add-in compatible with an existing COM add-in</span></span>](../../develop/make-office-add-in-compatible-with-existing-com-add-in.md)