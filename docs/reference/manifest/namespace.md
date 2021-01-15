---
title: Elemento Namespace no arquivo de manifesto
description: O elemento namespace define o namespace que uma função personalizada usa no Excel.
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: 342f5ebcafa861838956f1033f8597cf05e60215
ms.sourcegitcommit: 2f75a37de349251bc0e0fc402c5ae6dc5c3b8b08
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 01/06/2021
ms.locfileid: "49771253"
---
# <a name="namespace-element"></a><span data-ttu-id="89119-103">Elemento Namespace</span><span class="sxs-lookup"><span data-stu-id="89119-103">Namespace element</span></span>

<span data-ttu-id="89119-104">Define o namespace usado por uma função personalizada no Excel.</span><span class="sxs-lookup"><span data-stu-id="89119-104">Defines the namespace used by a custom function in Excel.</span></span>

## <a name="attributes"></a><span data-ttu-id="89119-105">Atributos</span><span class="sxs-lookup"><span data-stu-id="89119-105">Attributes</span></span>

|  <span data-ttu-id="89119-106">Atributo</span><span class="sxs-lookup"><span data-stu-id="89119-106">Attribute</span></span>  |  <span data-ttu-id="89119-107">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="89119-107">Required</span></span>  |  <span data-ttu-id="89119-108">Descrição</span><span class="sxs-lookup"><span data-stu-id="89119-108">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="89119-109">**resid="namespace"**</span><span class="sxs-lookup"><span data-stu-id="89119-109">**resid="namespace"**</span></span>  |  <span data-ttu-id="89119-110">Não</span><span class="sxs-lookup"><span data-stu-id="89119-110">No</span></span>  | <span data-ttu-id="89119-111">Deve corresponder ao título ShortStrings para sua função personalizada, especificada no elemento [Resources](resources.md).</span><span class="sxs-lookup"><span data-stu-id="89119-111">Should match the ShortStrings title for your custom function, specified within the [Resources](resources.md) element.</span></span> <span data-ttu-id="89119-112">Não pode ter mais de 32 caracteres.</span><span class="sxs-lookup"><span data-stu-id="89119-112">Can be no more than 32 characters.</span></span> |

## <a name="child-elements"></a><span data-ttu-id="89119-113">Elementos filho</span><span class="sxs-lookup"><span data-stu-id="89119-113">Child elements</span></span>

<span data-ttu-id="89119-114">Nenhum</span><span class="sxs-lookup"><span data-stu-id="89119-114">None</span></span>

## <a name="example"></a><span data-ttu-id="89119-115">Exemplo</span><span class="sxs-lookup"><span data-stu-id="89119-115">Example</span></span>

```xml
<Namespace resid="namespace" />
```
