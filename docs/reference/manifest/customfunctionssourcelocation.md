---
title: Elemento SourceLocation no arquivo de manifesto
description: ''
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: b2b78065fc8bde6fc827ddcb21e2bc700ed5bf49
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 04/24/2019
ms.locfileid: "32450685"
---
# <a name="sourcelocation-element"></a><span data-ttu-id="6a20f-102">Elemento SourceLocation</span><span class="sxs-lookup"><span data-stu-id="6a20f-102">SourceLocation element</span></span>

<span data-ttu-id="6a20f-103">Define a localização de um recurso necessário para os elementos de Página ou Script usados por funções personalizadas no Excel.</span><span class="sxs-lookup"><span data-stu-id="6a20f-103">Defines the location of a resource needed by the Script or Page elements used by custom functions in Excel.</span></span>

## <a name="attributes"></a><span data-ttu-id="6a20f-104">Atributos</span><span class="sxs-lookup"><span data-stu-id="6a20f-104">Attributes</span></span>

| <span data-ttu-id="6a20f-105">**Atributo**</span><span class="sxs-lookup"><span data-stu-id="6a20f-105">**Attribute**</span></span> | <span data-ttu-id="6a20f-106">**Obrigatório**</span><span class="sxs-lookup"><span data-stu-id="6a20f-106">**Required**</span></span> | <span data-ttu-id="6a20f-107">**Descrição**</span><span class="sxs-lookup"><span data-stu-id="6a20f-107">**Description**</span></span>                                                                      |
|---------------|--------------|--------------------------------------------------------------------------------------|
| <span data-ttu-id="6a20f-108">resid</span><span class="sxs-lookup"><span data-stu-id="6a20f-108">resid</span></span>         | <span data-ttu-id="6a20f-109">Sim</span><span class="sxs-lookup"><span data-stu-id="6a20f-109">Yes</span></span>          | <span data-ttu-id="6a20f-110">O nome de um recurso de URL definido na seção &lt;Recursos&gt; do manifesto.</span><span class="sxs-lookup"><span data-stu-id="6a20f-110">The name of a URL resource defined in the &lt;Resources&gt; section of the manifest.</span></span> |

## <a name="child-elements"></a><span data-ttu-id="6a20f-111">Elementos filho</span><span class="sxs-lookup"><span data-stu-id="6a20f-111">Child elements</span></span>

<span data-ttu-id="6a20f-112">Nenhum</span><span class="sxs-lookup"><span data-stu-id="6a20f-112">None</span></span>

## <a name="example"></a><span data-ttu-id="6a20f-113">Exemplo</span><span class="sxs-lookup"><span data-stu-id="6a20f-113">Example</span></span>

```xml
<SourceLocation resid="pageURL"/>
```
