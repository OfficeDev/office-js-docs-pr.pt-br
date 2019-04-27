---
title: Elemento Type no arquivo de manifesto
description: ''
ms.date: 04/22/2019
localization_priority: Normal
ms.openlocfilehash: 28514e25d7877c0452fbf006a31f078cd980d819
ms.sourcegitcommit: 7462409209264dc7f8f89f3808a7a6249fcd739e
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 04/26/2019
ms.locfileid: "33356841"
---
# <a name="type-element"></a><span data-ttu-id="85c0d-102">Elemento Type</span><span class="sxs-lookup"><span data-stu-id="85c0d-102">Type element</span></span>

<span data-ttu-id="85c0d-103">Especifica se o suplemento equivalente é um suplemento COM ou um XLL.</span><span class="sxs-lookup"><span data-stu-id="85c0d-103">Specifies if the equivalent add-in is a COM addin or an XLL.</span></span>

<span data-ttu-id="85c0d-104">**Tipo de suplemento:** Painel de tarefas, função personalizada</span><span class="sxs-lookup"><span data-stu-id="85c0d-104">**Add-in type:** Task pane, Custom function</span></span>

## <a name="syntax"></a><span data-ttu-id="85c0d-105">Sintaxe</span><span class="sxs-lookup"><span data-stu-id="85c0d-105">Syntax</span></span>

```XML
    <Type> [COM | XLL] </Type>  
```

## <a name="contained-in"></a><span data-ttu-id="85c0d-106">Contido em</span><span class="sxs-lookup"><span data-stu-id="85c0d-106">Contained in</span></span>

[<span data-ttu-id="85c0d-107">EquivalentAdd-in</span><span class="sxs-lookup"><span data-stu-id="85c0d-107">EquivalentAdd-in</span></span>](equivalentaddin.md)

## <a name="add-in-type-values"></a><span data-ttu-id="85c0d-108">Valores de tipo de suplemento</span><span class="sxs-lookup"><span data-stu-id="85c0d-108">Add-in type values</span></span>

<span data-ttu-id="85c0d-109">Você deve especificar um dos seguintes valores para o `Type` elemento.</span><span class="sxs-lookup"><span data-stu-id="85c0d-109">You must specify one of the following values for the `Type` element.</span></span>

- <span data-ttu-id="85c0d-110">COM: especifica o suplemento equivalente é um suplemento de COM.</span><span class="sxs-lookup"><span data-stu-id="85c0d-110">COM: Specifies the equivalent add-in is a COM add-in.</span></span>
- <span data-ttu-id="85c0d-111">XLL: especifica o suplemento equivalente é um XLL do Excel.</span><span class="sxs-lookup"><span data-stu-id="85c0d-111">XLL: Specifies the equivalent add-in is an Excel XLL.</span></span>

## <a name="see-also"></a><span data-ttu-id="85c0d-112">Confira também</span><span class="sxs-lookup"><span data-stu-id="85c0d-112">See also</span></span>

- [<span data-ttu-id="85c0d-113">Tornar suas funções personalizadas compatíveis com as funções definidas pelo usuário XLL</span><span class="sxs-lookup"><span data-stu-id="85c0d-113">Make your custom functions compatible with XLL user-defined functions</span></span>](../../excel/make-custom-functions-compatible-with-xll-udf.md)
- [<span data-ttu-id="85c0d-114">Tornar o suplemento do Office compatível com um suplemento de COM existente</span><span class="sxs-lookup"><span data-stu-id="85c0d-114">Make your Office Add-in compatible with an existing COM add-in</span></span>](../../develop/make-office-add-in-compatible-with-existing-com-add-in.md)