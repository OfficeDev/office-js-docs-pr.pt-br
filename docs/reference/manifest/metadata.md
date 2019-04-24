---
title: Elemento Metadata no arquivo de manifesto
description: ''
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: a3aecb1983905658f3a55fdb8bf0629a8d5ef474
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 04/24/2019
ms.locfileid: "32452043"
---
# <a name="metadata-element"></a><span data-ttu-id="3ad17-102">Elemento Metadata</span><span class="sxs-lookup"><span data-stu-id="3ad17-102">Metadata element</span></span>

<span data-ttu-id="3ad17-103">Define as configurações de metadados usados por uma função personalizada no Excel.</span><span class="sxs-lookup"><span data-stu-id="3ad17-103">Defines the metadata settings used by a custom function in Excel.</span></span>

## <a name="attributes"></a><span data-ttu-id="3ad17-104">Atributos</span><span class="sxs-lookup"><span data-stu-id="3ad17-104">Attributes</span></span>

<span data-ttu-id="3ad17-105">Nenhuma</span><span class="sxs-lookup"><span data-stu-id="3ad17-105">None</span></span>

## <a name="child-elements"></a><span data-ttu-id="3ad17-106">Elementos filho</span><span class="sxs-lookup"><span data-stu-id="3ad17-106">Child elements</span></span>

|  <span data-ttu-id="3ad17-107">Elemento</span><span class="sxs-lookup"><span data-stu-id="3ad17-107">Element</span></span>  |  <span data-ttu-id="3ad17-108">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="3ad17-108">Required</span></span>  |  <span data-ttu-id="3ad17-109">Descrição</span><span class="sxs-lookup"><span data-stu-id="3ad17-109">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="3ad17-110">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="3ad17-110">SourceLocation</span></span>](customfunctionssourcelocation.md)  |  <span data-ttu-id="3ad17-111">Sim</span><span class="sxs-lookup"><span data-stu-id="3ad17-111">Yes</span></span>  | <span data-ttu-id="3ad17-112">Cadeia de caracteres com a ID de recurso do arquivo JSON usado por funções personalizadas.</span><span class="sxs-lookup"><span data-stu-id="3ad17-112">String with the resource id of the JSON file used by custom functions.</span></span> |

## <a name="example"></a><span data-ttu-id="3ad17-113">Exemplo</span><span class="sxs-lookup"><span data-stu-id="3ad17-113">Example</span></span>

```xml
<Metadata>
    <SourceLocation resid="JSON-URL" />
</Metadata>
```
