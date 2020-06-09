---
title: Elemento Namespace no arquivo de manifesto
description: O elemento namespace define o namespace que uma função personalizada usa no Excel.
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: f4b3510c6c137bd303af8a3eaac8ebe66c5f4dc7
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 06/08/2020
ms.locfileid: "44612231"
---
# <a name="namespace-element"></a><span data-ttu-id="6f92a-103">Elemento Namespace</span><span class="sxs-lookup"><span data-stu-id="6f92a-103">Namespace element</span></span>

<span data-ttu-id="6f92a-104">Define o namespace usado por uma função personalizada no Excel.</span><span class="sxs-lookup"><span data-stu-id="6f92a-104">Defines the namespace used by a custom function in Excel.</span></span>

## <a name="attributes"></a><span data-ttu-id="6f92a-105">Atributos</span><span class="sxs-lookup"><span data-stu-id="6f92a-105">Attributes</span></span>

|  <span data-ttu-id="6f92a-106">Atributo</span><span class="sxs-lookup"><span data-stu-id="6f92a-106">Attribute</span></span>  |  <span data-ttu-id="6f92a-107">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="6f92a-107">Required</span></span>  |  <span data-ttu-id="6f92a-108">Descrição</span><span class="sxs-lookup"><span data-stu-id="6f92a-108">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="6f92a-109">**resid="namespace"**</span><span class="sxs-lookup"><span data-stu-id="6f92a-109">**resid="namespace"**</span></span>  |  <span data-ttu-id="6f92a-110">Não</span><span class="sxs-lookup"><span data-stu-id="6f92a-110">No</span></span>  | <span data-ttu-id="6f92a-111">Deve corresponder ao título ShortStrings para sua função personalizada, especificada no elemento [Resources](resources.md).</span><span class="sxs-lookup"><span data-stu-id="6f92a-111">Should match the ShortStrings title for your custom function, specified within the [Resources](resources.md) element.</span></span> |

## <a name="child-elements"></a><span data-ttu-id="6f92a-112">Elementos filho</span><span class="sxs-lookup"><span data-stu-id="6f92a-112">Child elements</span></span>

<span data-ttu-id="6f92a-113">Nenhum</span><span class="sxs-lookup"><span data-stu-id="6f92a-113">None</span></span>

## <a name="example"></a><span data-ttu-id="6f92a-114">Exemplo</span><span class="sxs-lookup"><span data-stu-id="6f92a-114">Example</span></span>

```xml
<Namespace resid="namespace" />
```
