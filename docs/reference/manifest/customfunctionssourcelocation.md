---
title: Elemento SourceLocation no arquivo de manifesto
description: Define a localização de um recurso necessário para os elementos de Página ou Script usados por funções personalizadas no Excel.
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: 56ebe122853c98a14c52d450bea31fecaefb15d3
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/17/2020
ms.locfileid: "42720684"
---
# <a name="sourcelocation-element"></a><span data-ttu-id="ebbf6-103">Elemento SourceLocation</span><span class="sxs-lookup"><span data-stu-id="ebbf6-103">SourceLocation element</span></span>

<span data-ttu-id="ebbf6-104">Define a localização de um recurso necessário para os elementos de Página ou Script usados por funções personalizadas no Excel.</span><span class="sxs-lookup"><span data-stu-id="ebbf6-104">Defines the location of a resource needed by the Script or Page elements used by custom functions in Excel.</span></span>

## <a name="attributes"></a><span data-ttu-id="ebbf6-105">Atributos</span><span class="sxs-lookup"><span data-stu-id="ebbf6-105">Attributes</span></span>

| <span data-ttu-id="ebbf6-106">**Atributo**</span><span class="sxs-lookup"><span data-stu-id="ebbf6-106">**Attribute**</span></span> | <span data-ttu-id="ebbf6-107">**Obrigatório**</span><span class="sxs-lookup"><span data-stu-id="ebbf6-107">**Required**</span></span> | <span data-ttu-id="ebbf6-108">**Descrição**</span><span class="sxs-lookup"><span data-stu-id="ebbf6-108">**Description**</span></span>                                                                      |
|---------------|--------------|--------------------------------------------------------------------------------------|
| <span data-ttu-id="ebbf6-109">resid</span><span class="sxs-lookup"><span data-stu-id="ebbf6-109">resid</span></span>         | <span data-ttu-id="ebbf6-110">Sim</span><span class="sxs-lookup"><span data-stu-id="ebbf6-110">Yes</span></span>          | <span data-ttu-id="ebbf6-111">O nome de um recurso de URL definido na seção &lt;Recursos&gt; do manifesto.</span><span class="sxs-lookup"><span data-stu-id="ebbf6-111">The name of a URL resource defined in the &lt;Resources&gt; section of the manifest.</span></span> |

## <a name="child-elements"></a><span data-ttu-id="ebbf6-112">Elementos filho</span><span class="sxs-lookup"><span data-stu-id="ebbf6-112">Child elements</span></span>

<span data-ttu-id="ebbf6-113">Nenhum</span><span class="sxs-lookup"><span data-stu-id="ebbf6-113">None</span></span>

## <a name="example"></a><span data-ttu-id="ebbf6-114">Exemplo</span><span class="sxs-lookup"><span data-stu-id="ebbf6-114">Example</span></span>

```xml
<SourceLocation resid="pageURL"/>
```
