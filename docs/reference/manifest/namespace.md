---
title: Elemento Namespace no arquivo de manifesto
description: O elemento namespace define o namespace que uma função personalizada usa no Excel.
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: 45fd0caa039fdeb885cba4b739750fbd8b642252
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/17/2020
ms.locfileid: "42718052"
---
# <a name="namespace-element"></a><span data-ttu-id="36090-103">Elemento Namespace</span><span class="sxs-lookup"><span data-stu-id="36090-103">Namespace element</span></span>

<span data-ttu-id="36090-104">Define o namespace usado por uma função personalizada no Excel.</span><span class="sxs-lookup"><span data-stu-id="36090-104">Defines the namespace used by a custom function in Excel.</span></span>

## <a name="attributes"></a><span data-ttu-id="36090-105">Atributos</span><span class="sxs-lookup"><span data-stu-id="36090-105">Attributes</span></span>

|  <span data-ttu-id="36090-106">Atributo</span><span class="sxs-lookup"><span data-stu-id="36090-106">Attribute</span></span>  |  <span data-ttu-id="36090-107">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="36090-107">Required</span></span>  |  <span data-ttu-id="36090-108">Descrição</span><span class="sxs-lookup"><span data-stu-id="36090-108">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="36090-109">**resid="namespace"**</span><span class="sxs-lookup"><span data-stu-id="36090-109">**resid="namespace"**</span></span>  |  <span data-ttu-id="36090-110">Sim</span><span class="sxs-lookup"><span data-stu-id="36090-110">Yes</span></span>  | <span data-ttu-id="36090-111">Deve corresponder ao título ShortStrings para sua função personalizada, especificada no elemento [Resources](resources.md).</span><span class="sxs-lookup"><span data-stu-id="36090-111">Should match the ShortStrings title for your custom function, specified within the [Resources](resources.md) element.</span></span> |

## <a name="child-elements"></a><span data-ttu-id="36090-112">Elementos filho</span><span class="sxs-lookup"><span data-stu-id="36090-112">Child elements</span></span>

<span data-ttu-id="36090-113">Nenhum</span><span class="sxs-lookup"><span data-stu-id="36090-113">None</span></span>

## <a name="example"></a><span data-ttu-id="36090-114">Exemplo</span><span class="sxs-lookup"><span data-stu-id="36090-114">Example</span></span>

```xml
<Namespace resid="namespace" />
```
