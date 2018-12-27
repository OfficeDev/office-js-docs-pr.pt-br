---
title: Elemento SourceLocation no arquivo de manifesto
description: ''
ms.date: 10/09/2018
ms.openlocfilehash: 878a8184984e31fdbcf46192a2f56507edaf4b37
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 12/22/2018
ms.locfileid: "27432402"
---
# <a name="sourcelocation-element"></a><span data-ttu-id="2c15b-102">Elemento SourceLocation</span><span class="sxs-lookup"><span data-stu-id="2c15b-102">SourceLocation element</span></span>

<span data-ttu-id="2c15b-103">Define a localização de um recurso necessário para os elementos de Página ou Script usados por funções personalizadas no Excel.</span><span class="sxs-lookup"><span data-stu-id="2c15b-103">Defines the location of a resource needed by the Script or Page elements used by custom functions in Excel.</span></span>

## <a name="attributes"></a><span data-ttu-id="2c15b-104">Atributos</span><span class="sxs-lookup"><span data-stu-id="2c15b-104">Attributes</span></span>

| <span data-ttu-id="2c15b-105">**Atributo**</span><span class="sxs-lookup"><span data-stu-id="2c15b-105">**Attribute**</span></span> | <span data-ttu-id="2c15b-106">**Obrigatório**</span><span class="sxs-lookup"><span data-stu-id="2c15b-106">**Required**</span></span> | <span data-ttu-id="2c15b-107">**Descrição**</span><span class="sxs-lookup"><span data-stu-id="2c15b-107">**Description**</span></span>                                                                      |
|---------------|--------------|--------------------------------------------------------------------------------------|
| <span data-ttu-id="2c15b-108">resid</span><span class="sxs-lookup"><span data-stu-id="2c15b-108">resid</span></span>         | <span data-ttu-id="2c15b-109">Sim</span><span class="sxs-lookup"><span data-stu-id="2c15b-109">Yes</span></span>          | <span data-ttu-id="2c15b-110">O nome de um recurso de URL definido na seção &lt;Recursos&gt; do manifesto.</span><span class="sxs-lookup"><span data-stu-id="2c15b-110">The name of a URL resource defined in the &lt;Resources&gt; section of the manifest.</span></span> |

## <a name="child-elements"></a><span data-ttu-id="2c15b-111">Elementos filho</span><span class="sxs-lookup"><span data-stu-id="2c15b-111">Child elements</span></span>

<span data-ttu-id="2c15b-112">Nenhum</span><span class="sxs-lookup"><span data-stu-id="2c15b-112">None</span></span>

## <a name="example"></a><span data-ttu-id="2c15b-113">Exemplo</span><span class="sxs-lookup"><span data-stu-id="2c15b-113">Example</span></span>

```xml
<SourceLocation resid="pageURL"/>
```