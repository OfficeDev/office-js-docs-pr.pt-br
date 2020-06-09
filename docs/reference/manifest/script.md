---
title: Elemento Script no arquivo de manifesto
description: O elemento script define as configurações de script que uma função personalizada usa no Excel.
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: 791f49f15673a029b982e40946f8cc90f02ba887
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 06/08/2020
ms.locfileid: "44608087"
---
# <a name="script-element"></a><span data-ttu-id="fdba9-103">Elemento Script</span><span class="sxs-lookup"><span data-stu-id="fdba9-103">Script element</span></span>

<span data-ttu-id="fdba9-104">Define as configurações de script usadas por uma função personalizada no Excel.</span><span class="sxs-lookup"><span data-stu-id="fdba9-104">Defines script settings used by a custom function in Excel.</span></span>

## <a name="attributes"></a><span data-ttu-id="fdba9-105">Atributos</span><span class="sxs-lookup"><span data-stu-id="fdba9-105">Attributes</span></span>

<span data-ttu-id="fdba9-106">Nenhuma</span><span class="sxs-lookup"><span data-stu-id="fdba9-106">None</span></span>

## <a name="child-elements"></a><span data-ttu-id="fdba9-107">Elementos filho</span><span class="sxs-lookup"><span data-stu-id="fdba9-107">Child elements</span></span>

|<span data-ttu-id="fdba9-108">Elementos</span><span class="sxs-lookup"><span data-stu-id="fdba9-108">Elements</span></span>  |  <span data-ttu-id="fdba9-109">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="fdba9-109">Required</span></span>  |  <span data-ttu-id="fdba9-110">Descrição</span><span class="sxs-lookup"><span data-stu-id="fdba9-110">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="fdba9-111">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="fdba9-111">SourceLocation</span></span>](customfunctionssourcelocation.md)  |  <span data-ttu-id="fdba9-112">Sim</span><span class="sxs-lookup"><span data-stu-id="fdba9-112">Yes</span></span>  | <span data-ttu-id="fdba9-113">Cadeia de caracteres com o ID de recurso do arquivo JavaScript usado por funções personalizadas.</span><span class="sxs-lookup"><span data-stu-id="fdba9-113">String with resource id of the JavaScript file used by custom functions.</span></span>|

## <a name="example"></a><span data-ttu-id="fdba9-114">Exemplo</span><span class="sxs-lookup"><span data-stu-id="fdba9-114">Example</span></span>

```xml
<Script>
    <SourceLocation resid="scriptURL" />
    <!-- The Script element is not used in the Developer Preview. -->
</Script>
```
