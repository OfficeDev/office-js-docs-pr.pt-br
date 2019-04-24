---
title: Elemento Script no arquivo de manifesto
description: ''
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: 8352ada0eeb6af071d5f20f750dcdeaefe31e918
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 04/24/2019
ms.locfileid: "32450433"
---
# <a name="script-element"></a><span data-ttu-id="453f6-102">Elemento Script</span><span class="sxs-lookup"><span data-stu-id="453f6-102">Script element</span></span>

<span data-ttu-id="453f6-103">Define as configurações de script usadas por uma função personalizada no Excel.</span><span class="sxs-lookup"><span data-stu-id="453f6-103">Defines script settings used by a custom function in Excel.</span></span>

## <a name="attributes"></a><span data-ttu-id="453f6-104">Atributos</span><span class="sxs-lookup"><span data-stu-id="453f6-104">Attributes</span></span>

<span data-ttu-id="453f6-105">Nenhuma</span><span class="sxs-lookup"><span data-stu-id="453f6-105">None</span></span>

## <a name="child-elements"></a><span data-ttu-id="453f6-106">Elementos filho</span><span class="sxs-lookup"><span data-stu-id="453f6-106">Child elements</span></span>

|<span data-ttu-id="453f6-107">Elementos</span><span class="sxs-lookup"><span data-stu-id="453f6-107">Elements</span></span>  |  <span data-ttu-id="453f6-108">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="453f6-108">Required</span></span>  |  <span data-ttu-id="453f6-109">Descrição</span><span class="sxs-lookup"><span data-stu-id="453f6-109">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="453f6-110">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="453f6-110">SourceLocation</span></span>](customfunctionssourcelocation.md)  |  <span data-ttu-id="453f6-111">Sim</span><span class="sxs-lookup"><span data-stu-id="453f6-111">Yes</span></span>  | <span data-ttu-id="453f6-112">Cadeia de caracteres com o ID de recurso do arquivo JavaScript usado por funções personalizadas.</span><span class="sxs-lookup"><span data-stu-id="453f6-112">String with resource id of the JavaScript file used by custom functions.</span></span>|

## <a name="example"></a><span data-ttu-id="453f6-113">Exemplo</span><span class="sxs-lookup"><span data-stu-id="453f6-113">Example</span></span>

```xml
<Script>
    <SourceLocation resid="scriptURL" />
    <!-- The Script element is not used in the Developer Preview. -->
</Script>
```
