---
title: Elemento Namespace no arquivo de manifesto
description: ''
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: faf77fe8b6bddc734f1b47eb544ffe7e1e7c4aaa
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 04/24/2019
ms.locfileid: "32452099"
---
# <a name="namespace-element"></a><span data-ttu-id="f42d5-102">Elemento Namespace</span><span class="sxs-lookup"><span data-stu-id="f42d5-102">Namespace element</span></span>

<span data-ttu-id="f42d5-103">Define o namespace usado por uma função personalizada no Excel.</span><span class="sxs-lookup"><span data-stu-id="f42d5-103">Defines the namespace used by a custom function in Excel.</span></span>

## <a name="attributes"></a><span data-ttu-id="f42d5-104">Atributos</span><span class="sxs-lookup"><span data-stu-id="f42d5-104">Attributes</span></span>

|  <span data-ttu-id="f42d5-105">Atributo</span><span class="sxs-lookup"><span data-stu-id="f42d5-105">Attribute</span></span>  |  <span data-ttu-id="f42d5-106">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="f42d5-106">Required</span></span>  |  <span data-ttu-id="f42d5-107">Descrição</span><span class="sxs-lookup"><span data-stu-id="f42d5-107">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="f42d5-108">**resid="namespace"**</span><span class="sxs-lookup"><span data-stu-id="f42d5-108">**resid="namespace"**</span></span>  |  <span data-ttu-id="f42d5-109">Sim</span><span class="sxs-lookup"><span data-stu-id="f42d5-109">Yes</span></span>  | <span data-ttu-id="f42d5-110">Deve corresponder ao título ShortStrings para sua função personalizada, especificada no elemento [Resources](resources.md).</span><span class="sxs-lookup"><span data-stu-id="f42d5-110">Should match the ShortStrings title for your custom function, specified within the [Resources](resources.md) element.</span></span> |

## <a name="child-elements"></a><span data-ttu-id="f42d5-111">Elementos filho</span><span class="sxs-lookup"><span data-stu-id="f42d5-111">Child elements</span></span>

<span data-ttu-id="f42d5-112">Nenhum</span><span class="sxs-lookup"><span data-stu-id="f42d5-112">None</span></span>

## <a name="example"></a><span data-ttu-id="f42d5-113">Exemplo</span><span class="sxs-lookup"><span data-stu-id="f42d5-113">Example</span></span>

```xml
<Namespace resid="namespace" />
```
