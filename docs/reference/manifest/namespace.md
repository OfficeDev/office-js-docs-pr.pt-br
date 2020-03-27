---
title: Elemento Namespace no arquivo de manifesto
description: O elemento namespace define o namespace que uma função personalizada usa no Excel.
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: eabd73d3be98271c81723787dd3d1bdb6ee2ebcd
ms.sourcegitcommit: 315a648cce38609c3e1c92bd4a339e268f8a2e1d
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/26/2020
ms.locfileid: "42978666"
---
# <a name="namespace-element"></a><span data-ttu-id="ea8aa-103">Elemento Namespace</span><span class="sxs-lookup"><span data-stu-id="ea8aa-103">Namespace element</span></span>

<span data-ttu-id="ea8aa-104">Define o namespace usado por uma função personalizada no Excel.</span><span class="sxs-lookup"><span data-stu-id="ea8aa-104">Defines the namespace used by a custom function in Excel.</span></span>

## <a name="attributes"></a><span data-ttu-id="ea8aa-105">Atributos</span><span class="sxs-lookup"><span data-stu-id="ea8aa-105">Attributes</span></span>

|  <span data-ttu-id="ea8aa-106">Atributo</span><span class="sxs-lookup"><span data-stu-id="ea8aa-106">Attribute</span></span>  |  <span data-ttu-id="ea8aa-107">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="ea8aa-107">Required</span></span>  |  <span data-ttu-id="ea8aa-108">Descrição</span><span class="sxs-lookup"><span data-stu-id="ea8aa-108">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="ea8aa-109">**resid="namespace"**</span><span class="sxs-lookup"><span data-stu-id="ea8aa-109">**resid="namespace"**</span></span>  |  <span data-ttu-id="ea8aa-110">Não</span><span class="sxs-lookup"><span data-stu-id="ea8aa-110">No</span></span>  | <span data-ttu-id="ea8aa-111">Deve corresponder ao título ShortStrings para sua função personalizada, especificada no elemento [Resources](resources.md).</span><span class="sxs-lookup"><span data-stu-id="ea8aa-111">Should match the ShortStrings title for your custom function, specified within the [Resources](resources.md) element.</span></span> |

## <a name="child-elements"></a><span data-ttu-id="ea8aa-112">Elementos filho</span><span class="sxs-lookup"><span data-stu-id="ea8aa-112">Child elements</span></span>

<span data-ttu-id="ea8aa-113">Nenhum</span><span class="sxs-lookup"><span data-stu-id="ea8aa-113">None</span></span>

## <a name="example"></a><span data-ttu-id="ea8aa-114">Exemplo</span><span class="sxs-lookup"><span data-stu-id="ea8aa-114">Example</span></span>

```xml
<Namespace resid="namespace" />
```
