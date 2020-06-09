---
title: Elemento Metadata no arquivo de manifesto
description: O elemento de metadados define as configurações de metadados que uma função personalizada usa no Excel.
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: 01be124b5526ce8328e0a20b8ff7d21ba6da96bc
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 06/08/2020
ms.locfileid: "44611761"
---
# <a name="metadata-element"></a><span data-ttu-id="e4dea-103">Elemento Metadata</span><span class="sxs-lookup"><span data-stu-id="e4dea-103">Metadata element</span></span>

<span data-ttu-id="e4dea-104">Define as configurações de metadados usados por uma função personalizada no Excel.</span><span class="sxs-lookup"><span data-stu-id="e4dea-104">Defines the metadata settings used by a custom function in Excel.</span></span>

## <a name="attributes"></a><span data-ttu-id="e4dea-105">Atributos</span><span class="sxs-lookup"><span data-stu-id="e4dea-105">Attributes</span></span>

<span data-ttu-id="e4dea-106">Nenhuma</span><span class="sxs-lookup"><span data-stu-id="e4dea-106">None</span></span>

## <a name="child-elements"></a><span data-ttu-id="e4dea-107">Elementos filho</span><span class="sxs-lookup"><span data-stu-id="e4dea-107">Child elements</span></span>

|  <span data-ttu-id="e4dea-108">Elemento</span><span class="sxs-lookup"><span data-stu-id="e4dea-108">Element</span></span>  |  <span data-ttu-id="e4dea-109">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="e4dea-109">Required</span></span>  |  <span data-ttu-id="e4dea-110">Descrição</span><span class="sxs-lookup"><span data-stu-id="e4dea-110">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="e4dea-111">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="e4dea-111">SourceLocation</span></span>](customfunctionssourcelocation.md)  |  <span data-ttu-id="e4dea-112">Sim</span><span class="sxs-lookup"><span data-stu-id="e4dea-112">Yes</span></span>  | <span data-ttu-id="e4dea-113">Cadeia de caracteres com a ID de recurso do arquivo JSON usado por funções personalizadas.</span><span class="sxs-lookup"><span data-stu-id="e4dea-113">String with the resource id of the JSON file used by custom functions.</span></span> |

## <a name="example"></a><span data-ttu-id="e4dea-114">Exemplo</span><span class="sxs-lookup"><span data-stu-id="e4dea-114">Example</span></span>

```xml
<Metadata>
    <SourceLocation resid="JSON-URL" />
</Metadata>
```
