---
title: Elemento Script no arquivo de manifesto
description: ''
ms.date: 10/09/2018
ms.openlocfilehash: 95e4cbadc35302b4f76108e0ff2a51d31ca89aac
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 12/22/2018
ms.locfileid: "27433135"
---
# <a name="script-element"></a><span data-ttu-id="33591-102">Elemento Script</span><span class="sxs-lookup"><span data-stu-id="33591-102">Script element</span></span>

<span data-ttu-id="33591-103">Define as configurações de script usadas por uma função personalizada no Excel.</span><span class="sxs-lookup"><span data-stu-id="33591-103">Defines script settings used by a custom function in Excel.</span></span>

## <a name="attributes"></a><span data-ttu-id="33591-104">Atributos</span><span class="sxs-lookup"><span data-stu-id="33591-104">Attributes</span></span>

<span data-ttu-id="33591-105">Nenhuma</span><span class="sxs-lookup"><span data-stu-id="33591-105">None</span></span>

## <a name="child-elements"></a><span data-ttu-id="33591-106">Elementos filho</span><span class="sxs-lookup"><span data-stu-id="33591-106">Child elements</span></span>

|<span data-ttu-id="33591-107">Elementos</span><span class="sxs-lookup"><span data-stu-id="33591-107">Elements</span></span>  |  <span data-ttu-id="33591-108">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="33591-108">Required</span></span>  |  <span data-ttu-id="33591-109">Descrição</span><span class="sxs-lookup"><span data-stu-id="33591-109">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="33591-110">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="33591-110">SourceLocation</span></span>](customfunctionssourcelocation.md)  |  <span data-ttu-id="33591-111">Sim</span><span class="sxs-lookup"><span data-stu-id="33591-111">Yes</span></span>  | <span data-ttu-id="33591-112">Cadeia de caracteres com o ID de recurso do arquivo JavaScript usado por funções personalizadas.</span><span class="sxs-lookup"><span data-stu-id="33591-112">String with resource id of the JavaScript file used by custom functions.</span></span>|

## <a name="example"></a><span data-ttu-id="33591-113">Exemplo</span><span class="sxs-lookup"><span data-stu-id="33591-113">Example</span></span>

```xml
<Script>
    <SourceLocation resid="scriptURL" />
    <!-- The Script element is not used in the Developer Preview. -->
</Script>
```
