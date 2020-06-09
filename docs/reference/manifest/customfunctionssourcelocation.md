---
title: Elemento SourceLocation no arquivo de manifesto
description: Define a localização de um recurso necessário para os elementos de Página ou Script usados por funções personalizadas no Excel.
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: 88ae0558577167074a870170833617c4f60730f1
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 06/08/2020
ms.locfileid: "44612309"
---
# <a name="sourcelocation-element"></a><span data-ttu-id="1bcc3-103">Elemento SourceLocation</span><span class="sxs-lookup"><span data-stu-id="1bcc3-103">SourceLocation element</span></span>

<span data-ttu-id="1bcc3-104">Define a localização de um recurso necessário para os elementos de Página ou Script usados por funções personalizadas no Excel.</span><span class="sxs-lookup"><span data-stu-id="1bcc3-104">Defines the location of a resource needed by the Script or Page elements used by custom functions in Excel.</span></span>

## <a name="attributes"></a><span data-ttu-id="1bcc3-105">Atributos</span><span class="sxs-lookup"><span data-stu-id="1bcc3-105">Attributes</span></span>

| <span data-ttu-id="1bcc3-106">**Atributo**</span><span class="sxs-lookup"><span data-stu-id="1bcc3-106">**Attribute**</span></span> | <span data-ttu-id="1bcc3-107">**Obrigatório**</span><span class="sxs-lookup"><span data-stu-id="1bcc3-107">**Required**</span></span> | <span data-ttu-id="1bcc3-108">**Descrição**</span><span class="sxs-lookup"><span data-stu-id="1bcc3-108">**Description**</span></span>                                                                      |
|---------------|--------------|--------------------------------------------------------------------------------------|
| <span data-ttu-id="1bcc3-109">resid</span><span class="sxs-lookup"><span data-stu-id="1bcc3-109">resid</span></span>         | <span data-ttu-id="1bcc3-110">Sim</span><span class="sxs-lookup"><span data-stu-id="1bcc3-110">Yes</span></span>          | <span data-ttu-id="1bcc3-111">O nome de um recurso de URL definido na seção &lt;Recursos&gt; do manifesto.</span><span class="sxs-lookup"><span data-stu-id="1bcc3-111">The name of a URL resource defined in the &lt;Resources&gt; section of the manifest.</span></span> |

## <a name="child-elements"></a><span data-ttu-id="1bcc3-112">Elementos filho</span><span class="sxs-lookup"><span data-stu-id="1bcc3-112">Child elements</span></span>

<span data-ttu-id="1bcc3-113">Nenhum</span><span class="sxs-lookup"><span data-stu-id="1bcc3-113">None</span></span>

## <a name="example"></a><span data-ttu-id="1bcc3-114">Exemplo</span><span class="sxs-lookup"><span data-stu-id="1bcc3-114">Example</span></span>

```xml
<SourceLocation resid="pageURL"/>
```
