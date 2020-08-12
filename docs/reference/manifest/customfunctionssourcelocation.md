---
title: Elemento SourceLocation para funções personalizadas no arquivo de manifesto
description: Define a localização de um recurso necessário para os elementos de Página ou Script usados por funções personalizadas no Excel.
ms.date: 08/07/2020
localization_priority: Normal
ms.openlocfilehash: 1c509987b0ce7948a63fa8ad51f7cf9c84144c5f
ms.sourcegitcommit: cc6886b47c84ac37a3c957ff85dd0ed526ca5e43
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 08/12/2020
ms.locfileid: "46641379"
---
# <a name="sourcelocation-element-custom-functions"></a><span data-ttu-id="e19da-103">Elemento SourceLocation (funções personalizadas)</span><span class="sxs-lookup"><span data-stu-id="e19da-103">SourceLocation element (custom functions)</span></span>

<span data-ttu-id="e19da-104">Define a localização de um recurso necessário para os elementos de Página ou Script usados por funções personalizadas no Excel.</span><span class="sxs-lookup"><span data-stu-id="e19da-104">Defines the location of a resource needed by the Script or Page elements used by custom functions in Excel.</span></span>

## <a name="attributes"></a><span data-ttu-id="e19da-105">Atributos</span><span class="sxs-lookup"><span data-stu-id="e19da-105">Attributes</span></span>

| <span data-ttu-id="e19da-106">Atributo</span><span class="sxs-lookup"><span data-stu-id="e19da-106">Attribute</span></span> | <span data-ttu-id="e19da-107">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="e19da-107">Required</span></span> | <span data-ttu-id="e19da-108">Descrição</span><span class="sxs-lookup"><span data-stu-id="e19da-108">Description</span></span>                                                                          |
|-----------|----------|--------------------------------------------------------------------------------------|
| <span data-ttu-id="e19da-109">resid</span><span class="sxs-lookup"><span data-stu-id="e19da-109">resid</span></span>     | <span data-ttu-id="e19da-110">Sim</span><span class="sxs-lookup"><span data-stu-id="e19da-110">Yes</span></span>      | <span data-ttu-id="e19da-111">O nome de um recurso de URL definido na seção &lt;Recursos&gt; do manifesto.</span><span class="sxs-lookup"><span data-stu-id="e19da-111">The name of a URL resource defined in the &lt;Resources&gt; section of the manifest.</span></span> |

## <a name="child-elements"></a><span data-ttu-id="e19da-112">Elementos filho</span><span class="sxs-lookup"><span data-stu-id="e19da-112">Child elements</span></span>

<span data-ttu-id="e19da-113">Nenhum</span><span class="sxs-lookup"><span data-stu-id="e19da-113">None</span></span>

## <a name="example"></a><span data-ttu-id="e19da-114">Exemplo</span><span class="sxs-lookup"><span data-stu-id="e19da-114">Example</span></span>

```xml
<SourceLocation resid="pageURL"/>
```
