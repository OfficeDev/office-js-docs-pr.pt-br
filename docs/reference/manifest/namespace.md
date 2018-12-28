---
title: Elemento Namespace no arquivo de manifesto
description: ''
ms.date: 10/09/2018
ms.openlocfilehash: 8000ea5774b38dd038888c686a33127a2d5bc482
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 12/22/2018
ms.locfileid: "27432323"
---
# <a name="namespace-element"></a><span data-ttu-id="18198-102">Elemento Namespace</span><span class="sxs-lookup"><span data-stu-id="18198-102">Namespace element</span></span>

<span data-ttu-id="18198-103">Define o namespace usado por uma função personalizada no Excel.</span><span class="sxs-lookup"><span data-stu-id="18198-103">Defines script settings used by a custom function in Excel.</span></span>

## <a name="attributes"></a><span data-ttu-id="18198-104">Atributos</span><span class="sxs-lookup"><span data-stu-id="18198-104">Attributes</span></span>

|  <span data-ttu-id="18198-105">Atributo</span><span class="sxs-lookup"><span data-stu-id="18198-105">Attribute</span></span>  |  <span data-ttu-id="18198-106">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="18198-106">Required</span></span>  |  <span data-ttu-id="18198-107">Descrição</span><span class="sxs-lookup"><span data-stu-id="18198-107">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="18198-108">**resid="namespace"**</span><span class="sxs-lookup"><span data-stu-id="18198-108">**resid="namespace"**</span></span>  |  <span data-ttu-id="18198-109">Sim</span><span class="sxs-lookup"><span data-stu-id="18198-109">Yes</span></span>  | <span data-ttu-id="18198-110">Deve corresponder ao título ShortStrings para sua função personalizada, especificada no elemento [Resources](resources.md).</span><span class="sxs-lookup"><span data-stu-id="18198-110">Should match the ShortStrings title for your custom function, specified within the [Resources](resources.md) element.</span></span> |

## <a name="child-elements"></a><span data-ttu-id="18198-111">Elementos filho</span><span class="sxs-lookup"><span data-stu-id="18198-111">Child elements</span></span>

<span data-ttu-id="18198-112">Nenhum</span><span class="sxs-lookup"><span data-stu-id="18198-112">None</span></span>

## <a name="example"></a><span data-ttu-id="18198-113">Exemplo</span><span class="sxs-lookup"><span data-stu-id="18198-113">Example</span></span>

```xml
<Namespace resid="namespace" />
```
