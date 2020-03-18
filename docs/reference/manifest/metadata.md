---
title: Elemento Metadata no arquivo de manifesto
description: O elemento de metadados define as configurações de metadados que uma função personalizada usa no Excel.
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: 8ea81818aa96b407ce386ec318495ec5ba773d05
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/17/2020
ms.locfileid: "42718066"
---
# <a name="metadata-element"></a><span data-ttu-id="36617-103">Elemento Metadata</span><span class="sxs-lookup"><span data-stu-id="36617-103">Metadata element</span></span>

<span data-ttu-id="36617-104">Define as configurações de metadados usados por uma função personalizada no Excel.</span><span class="sxs-lookup"><span data-stu-id="36617-104">Defines the metadata settings used by a custom function in Excel.</span></span>

## <a name="attributes"></a><span data-ttu-id="36617-105">Atributos</span><span class="sxs-lookup"><span data-stu-id="36617-105">Attributes</span></span>

<span data-ttu-id="36617-106">Nenhuma</span><span class="sxs-lookup"><span data-stu-id="36617-106">None</span></span>

## <a name="child-elements"></a><span data-ttu-id="36617-107">Elementos filho</span><span class="sxs-lookup"><span data-stu-id="36617-107">Child elements</span></span>

|  <span data-ttu-id="36617-108">Elemento</span><span class="sxs-lookup"><span data-stu-id="36617-108">Element</span></span>  |  <span data-ttu-id="36617-109">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="36617-109">Required</span></span>  |  <span data-ttu-id="36617-110">Descrição</span><span class="sxs-lookup"><span data-stu-id="36617-110">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="36617-111">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="36617-111">SourceLocation</span></span>](customfunctionssourcelocation.md)  |  <span data-ttu-id="36617-112">Sim</span><span class="sxs-lookup"><span data-stu-id="36617-112">Yes</span></span>  | <span data-ttu-id="36617-113">Cadeia de caracteres com a ID de recurso do arquivo JSON usado por funções personalizadas.</span><span class="sxs-lookup"><span data-stu-id="36617-113">String with the resource id of the JSON file used by custom functions.</span></span> |

## <a name="example"></a><span data-ttu-id="36617-114">Exemplo</span><span class="sxs-lookup"><span data-stu-id="36617-114">Example</span></span>

```xml
<Metadata>
    <SourceLocation resid="JSON-URL" />
</Metadata>
```
