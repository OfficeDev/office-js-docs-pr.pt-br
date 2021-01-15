---
title: Elemento SourceLocation para funções personalizadas no arquivo de manifesto
description: Define a localização de um recurso necessário para os elementos de Página ou Script usados por funções personalizadas no Excel.
ms.date: 08/07/2020
localization_priority: Normal
ms.openlocfilehash: 6001673f1954a4af2de66ff7611069c3fb402a13
ms.sourcegitcommit: 2f75a37de349251bc0e0fc402c5ae6dc5c3b8b08
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 01/06/2021
ms.locfileid: "49771379"
---
# <a name="sourcelocation-element-custom-functions"></a><span data-ttu-id="8a146-103">Elemento SourceLocation (funções personalizadas)</span><span class="sxs-lookup"><span data-stu-id="8a146-103">SourceLocation element (custom functions)</span></span>

<span data-ttu-id="8a146-104">Define a localização de um recurso necessário para os elementos de Página ou Script usados por funções personalizadas no Excel.</span><span class="sxs-lookup"><span data-stu-id="8a146-104">Defines the location of a resource needed by the Script or Page elements used by custom functions in Excel.</span></span>

## <a name="attributes"></a><span data-ttu-id="8a146-105">Atributos</span><span class="sxs-lookup"><span data-stu-id="8a146-105">Attributes</span></span>

| <span data-ttu-id="8a146-106">Atributo</span><span class="sxs-lookup"><span data-stu-id="8a146-106">Attribute</span></span> | <span data-ttu-id="8a146-107">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="8a146-107">Required</span></span> | <span data-ttu-id="8a146-108">Descrição</span><span class="sxs-lookup"><span data-stu-id="8a146-108">Description</span></span>                                                                          |
|-----------|----------|--------------------------------------------------------------------------------------|
| <span data-ttu-id="8a146-109">resid</span><span class="sxs-lookup"><span data-stu-id="8a146-109">resid</span></span>     | <span data-ttu-id="8a146-110">Sim</span><span class="sxs-lookup"><span data-stu-id="8a146-110">Yes</span></span>      | <span data-ttu-id="8a146-111">O nome de um recurso de URL definido na seção &lt;Recursos&gt; do manifesto.</span><span class="sxs-lookup"><span data-stu-id="8a146-111">The name of a URL resource defined in the &lt;Resources&gt; section of the manifest.</span></span> <span data-ttu-id="8a146-112">Não pode ter mais de 32 caracteres.</span><span class="sxs-lookup"><span data-stu-id="8a146-112">Can be no more than 32 characters.</span></span> |

## <a name="child-elements"></a><span data-ttu-id="8a146-113">Elementos filho</span><span class="sxs-lookup"><span data-stu-id="8a146-113">Child elements</span></span>

<span data-ttu-id="8a146-114">Nenhum</span><span class="sxs-lookup"><span data-stu-id="8a146-114">None</span></span>

## <a name="example"></a><span data-ttu-id="8a146-115">Exemplo</span><span class="sxs-lookup"><span data-stu-id="8a146-115">Example</span></span>

```xml
<SourceLocation resid="pageURL"/>
```
