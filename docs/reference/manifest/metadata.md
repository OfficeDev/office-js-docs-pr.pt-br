---
title: Elemento Metadata no arquivo de manifesto
description: ''
ms.date: 10/09/2018
ms.openlocfilehash: 79038fc13eba76176be19e484ffa57e64727bf94
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 12/22/2018
ms.locfileid: "27432659"
---
# <a name="metadata-element"></a><span data-ttu-id="8bf41-102">Elemento Metadata</span><span class="sxs-lookup"><span data-stu-id="8bf41-102">MetaData element</span></span>

<span data-ttu-id="8bf41-103">Define as configurações de metadados usados por uma função personalizada no Excel.</span><span class="sxs-lookup"><span data-stu-id="8bf41-103">Defines script settings used by a custom function in Excel.</span></span>

## <a name="attributes"></a><span data-ttu-id="8bf41-104">Atributos</span><span class="sxs-lookup"><span data-stu-id="8bf41-104">Attributes</span></span>

<span data-ttu-id="8bf41-105">Nenhuma</span><span class="sxs-lookup"><span data-stu-id="8bf41-105">None</span></span>

## <a name="child-elements"></a><span data-ttu-id="8bf41-106">Elementos filho</span><span class="sxs-lookup"><span data-stu-id="8bf41-106">Child elements</span></span>

|  <span data-ttu-id="8bf41-107">Elemento</span><span class="sxs-lookup"><span data-stu-id="8bf41-107">Element</span></span>  |  <span data-ttu-id="8bf41-108">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="8bf41-108">Required</span></span>  |  <span data-ttu-id="8bf41-109">Descrição</span><span class="sxs-lookup"><span data-stu-id="8bf41-109">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="8bf41-110">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="8bf41-110">SourceLocation</span></span>](customfunctionssourcelocation.md)  |  <span data-ttu-id="8bf41-111">Sim</span><span class="sxs-lookup"><span data-stu-id="8bf41-111">Yes</span></span>  | <span data-ttu-id="8bf41-112">Cadeia de caracteres com a ID de recurso do arquivo JSON usado por funções personalizadas.</span><span class="sxs-lookup"><span data-stu-id="8bf41-112">String with the resource id of the HTML file used by custom functions.</span></span> |

## <a name="example"></a><span data-ttu-id="8bf41-113">Exemplo</span><span class="sxs-lookup"><span data-stu-id="8bf41-113">Example</span></span>

```xml
<Metadata>
    <SourceLocation resid="JSON-URL" />
</Metadata>
```
