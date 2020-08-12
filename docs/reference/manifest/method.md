---
title: Elemento Method no arquivo de manifesto
description: O elemento Method especifica um método individual da API JavaScript do Office que seus suplementos do Office exigem para ativar.
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: 0e3e74a73a3422a7789e82d6f0e7a516bd795ca8
ms.sourcegitcommit: cc6886b47c84ac37a3c957ff85dd0ed526ca5e43
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 08/12/2020
ms.locfileid: "46641322"
---
# <a name="method-element"></a><span data-ttu-id="7fe3f-103">Elemento Method</span><span class="sxs-lookup"><span data-stu-id="7fe3f-103">Method element</span></span>

<span data-ttu-id="7fe3f-104">Especifica um método individual da API JavaScript do Office que seu suplemento do Office exige para ativar.</span><span class="sxs-lookup"><span data-stu-id="7fe3f-104">Specifies an individual method from the Office JavaScript API that your Office Add-in requires in order to activate.</span></span>

<span data-ttu-id="7fe3f-105">**Tipo de suplemento:** Conteúdo, Painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="7fe3f-105">**Add-in type:** Content, Task pane</span></span>

## <a name="syntax"></a><span data-ttu-id="7fe3f-106">Sintaxe</span><span class="sxs-lookup"><span data-stu-id="7fe3f-106">Syntax</span></span>

```XML
<Method Name="string"/>
```

## <a name="contained-in"></a><span data-ttu-id="7fe3f-107">Contido em</span><span class="sxs-lookup"><span data-stu-id="7fe3f-107">Contained in</span></span>

[<span data-ttu-id="7fe3f-108">Methods</span><span class="sxs-lookup"><span data-stu-id="7fe3f-108">Methods</span></span>](methods.md)

## <a name="attributes"></a><span data-ttu-id="7fe3f-109">Atributos</span><span class="sxs-lookup"><span data-stu-id="7fe3f-109">Attributes</span></span>

|<span data-ttu-id="7fe3f-110">Atributo</span><span class="sxs-lookup"><span data-stu-id="7fe3f-110">Attribute</span></span>|<span data-ttu-id="7fe3f-111">Tipo</span><span class="sxs-lookup"><span data-stu-id="7fe3f-111">Type</span></span>|<span data-ttu-id="7fe3f-112">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="7fe3f-112">Required</span></span>|<span data-ttu-id="7fe3f-113">Descrição</span><span class="sxs-lookup"><span data-stu-id="7fe3f-113">Description</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="7fe3f-114">Nome</span><span class="sxs-lookup"><span data-stu-id="7fe3f-114">Name</span></span>|<span data-ttu-id="7fe3f-115">cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="7fe3f-115">string</span></span>|<span data-ttu-id="7fe3f-116">obrigatório</span><span class="sxs-lookup"><span data-stu-id="7fe3f-116">required</span></span>|<span data-ttu-id="7fe3f-117">Especifica o nome do método necessário qualificado com seu objeto pai.</span><span class="sxs-lookup"><span data-stu-id="7fe3f-117">Specifies the name of the required method qualified with its parent object.</span></span> <span data-ttu-id="7fe3f-118">Por exemplo, para especificar o `getSelectedDataAsync` método, você deve especificar `"Document.getSelectedDataAsync"` .</span><span class="sxs-lookup"><span data-stu-id="7fe3f-118">For example, to specify the `getSelectedDataAsync` method, you must specify `"Document.getSelectedDataAsync"`.</span></span>|

## <a name="remarks"></a><span data-ttu-id="7fe3f-119">Comentários</span><span class="sxs-lookup"><span data-stu-id="7fe3f-119">Remarks</span></span>

<span data-ttu-id="7fe3f-120">Os `Methods` `Method` elementos e não são suportados por suplementos de email. Para obter mais informações sobre conjuntos de requisitos, confira [versões do Office e conjuntos de requisitos](../../develop/office-versions-and-requirement-sets.md).</span><span class="sxs-lookup"><span data-stu-id="7fe3f-120">The `Methods` and `Method` elements aren't supported by mail add-ins. For more information about requirement sets, see [Office versions and requirement sets](../../develop/office-versions-and-requirement-sets.md).</span></span>

> [!IMPORTANT]
> <span data-ttu-id="7fe3f-121">Como não há forma de especificar o requisito de versão mínimo de métodos individuais, para verificar se um método está disponível no tempo de execução, você também deve usar uma instrução **if** ao chamar esse método no script do suplemento.</span><span class="sxs-lookup"><span data-stu-id="7fe3f-121">Because there is no way to specify the minimum version requirement for individual methods, to make sure that a method is available at runtime, you should also use an **if** statement when calling that method in the script of your add-in.</span></span> <span data-ttu-id="7fe3f-122">Para obter mais informações sobre como fazer isso, consulte [Understanding the Office JavaScript API](../../develop/understanding-the-javascript-api-for-office.md).</span><span class="sxs-lookup"><span data-stu-id="7fe3f-122">For more information about how to do this, see [Understanding the Office JavaScript API](../../develop/understanding-the-javascript-api-for-office.md).</span></span>
