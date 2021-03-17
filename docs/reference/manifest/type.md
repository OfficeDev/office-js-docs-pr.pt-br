---
title: Elemento Type no arquivo de manifesto
description: O elemento Type especifica se o complemento equivalente é um complemento COM ou um XLL.
ms.date: 03/09/2021
localization_priority: Normal
ms.openlocfilehash: 5af3359c232e91b097311bfc06fc9b1c932b0703
ms.sourcegitcommit: c0c61fe84f3c5de88bd7eac29120056bb1224fc8
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/17/2021
ms.locfileid: "50836806"
---
# <a name="type-element"></a><span data-ttu-id="a2899-103">Elemento Type</span><span class="sxs-lookup"><span data-stu-id="a2899-103">Type element</span></span>

<span data-ttu-id="a2899-104">Especifica se o complemento equivalente é um complemento COM ou um XLL.</span><span class="sxs-lookup"><span data-stu-id="a2899-104">Specifies if the equivalent add-in is a COM add-in or an XLL.</span></span>

<span data-ttu-id="a2899-105">**Tipo de complemento:** Painel de tarefas, função Personalizada</span><span class="sxs-lookup"><span data-stu-id="a2899-105">**Add-in type:** Task pane, Custom function</span></span>

## <a name="syntax"></a><span data-ttu-id="a2899-106">Sintaxe</span><span class="sxs-lookup"><span data-stu-id="a2899-106">Syntax</span></span>

```XML
    <Type> [COM | XLL] </Type>  
```

## <a name="contained-in"></a><span data-ttu-id="a2899-107">Contido em</span><span class="sxs-lookup"><span data-stu-id="a2899-107">Contained in</span></span>

[<span data-ttu-id="a2899-108">EquivalentAddin</span><span class="sxs-lookup"><span data-stu-id="a2899-108">EquivalentAddin</span></span>](equivalentaddin.md)

## <a name="add-in-type-values"></a><span data-ttu-id="a2899-109">Valores de tipo de complemento</span><span class="sxs-lookup"><span data-stu-id="a2899-109">Add-in type values</span></span>

<span data-ttu-id="a2899-110">Você deve especificar um dos seguintes valores para o `Type` elemento.</span><span class="sxs-lookup"><span data-stu-id="a2899-110">You must specify one of the following values for the `Type` element.</span></span>

- <span data-ttu-id="a2899-111">COM: Especifica que o complemento equivalente é um complemento COM.</span><span class="sxs-lookup"><span data-stu-id="a2899-111">COM: Specifies the equivalent add-in is a COM add-in.</span></span>
- <span data-ttu-id="a2899-112">XLL: Especifica que o complemento equivalente é um XLL do Excel.</span><span class="sxs-lookup"><span data-stu-id="a2899-112">XLL: Specifies the equivalent add-in is an Excel XLL.</span></span>

## <a name="see-also"></a><span data-ttu-id="a2899-113">Confira também</span><span class="sxs-lookup"><span data-stu-id="a2899-113">See also</span></span>

- [<span data-ttu-id="a2899-114">Torne as suas funções personalizadas compatíveis com as funções XLL definidas pelo usuário</span><span class="sxs-lookup"><span data-stu-id="a2899-114">Make your custom functions compatible with XLL user-defined functions</span></span>](../../excel/make-custom-functions-compatible-with-xll-udf.md)
- [<span data-ttu-id="a2899-115">Torne o seu suplemento do Office compatível com um suplemento COM existente</span><span class="sxs-lookup"><span data-stu-id="a2899-115">Make your Office Add-in compatible with an existing COM add-in</span></span>](../../develop/make-office-add-in-compatible-with-existing-com-add-in.md)