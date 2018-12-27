---
title: Elemento Page no arquivo de manifesto
description: ''
ms.date: 10/09/2018
ms.openlocfilehash: 83bafd24d0b56322ea5f7d51025f2416be019168
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 12/22/2018
ms.locfileid: "27433729"
---
# <a name="page-element"></a><span data-ttu-id="be8e0-102">Elemento Page</span><span class="sxs-lookup"><span data-stu-id="be8e0-102">Page element</span></span>

<span data-ttu-id="be8e0-103">Define as configurações de página HTML usadas por uma função personalizada no Excel.</span><span class="sxs-lookup"><span data-stu-id="be8e0-103">Defines HTML page settings used by a custom function in Excel.</span></span>

## <a name="attributes"></a><span data-ttu-id="be8e0-104">Atributos</span><span class="sxs-lookup"><span data-stu-id="be8e0-104">Attributes</span></span>

<span data-ttu-id="be8e0-105">Nenhuma</span><span class="sxs-lookup"><span data-stu-id="be8e0-105">None</span></span>

## <a name="child-elements"></a><span data-ttu-id="be8e0-106">Elementos filho</span><span class="sxs-lookup"><span data-stu-id="be8e0-106">Child elements</span></span>

|  <span data-ttu-id="be8e0-107">Elemento</span><span class="sxs-lookup"><span data-stu-id="be8e0-107">Element</span></span>  |  <span data-ttu-id="be8e0-108">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="be8e0-108">Required</span></span>  |  <span data-ttu-id="be8e0-109">Descrição</span><span class="sxs-lookup"><span data-stu-id="be8e0-109">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="be8e0-110">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="be8e0-110">SourceLocation</span></span>](customfunctionssourcelocation.md)  |  <span data-ttu-id="be8e0-111">Sim</span><span class="sxs-lookup"><span data-stu-id="be8e0-111">Yes</span></span>  | <span data-ttu-id="be8e0-112">Cadeia de caracteres com o ID de recurso do arquivo HTML usado por funções personalizadas.</span><span class="sxs-lookup"><span data-stu-id="be8e0-112">String with the resource id of the HTML file used by custom functions.</span></span> |

## <a name="example"></a><span data-ttu-id="be8e0-113">Exemplo</span><span class="sxs-lookup"><span data-stu-id="be8e0-113">Example</span></span>

```xml
<Page>
    <SourceLocation resid="pageURL"/>
</Page>
```
