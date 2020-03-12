---
title: Elemento Method no arquivo de manifesto
description: ''
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: 74b7a8b3d0f8511d21eb0df150500850e8b93fe9
ms.sourcegitcommit: 4079903c3cc45b7d8c041509a44e9fc38da399b1
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/11/2020
ms.locfileid: "42596890"
---
# <a name="method-element"></a><span data-ttu-id="f6bd8-102">Elemento Method</span><span class="sxs-lookup"><span data-stu-id="f6bd8-102">Method element</span></span>

<span data-ttu-id="f6bd8-103">Especifica um método individual da API JavaScript do Office que seu suplemento do Office exige para ativar.</span><span class="sxs-lookup"><span data-stu-id="f6bd8-103">Specifies an individual method from the Office JavaScript API that your Office Add-in requires in order to activate.</span></span>

<span data-ttu-id="f6bd8-104">**Tipo de suplemento:** Conteúdo, Painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="f6bd8-104">**Add-in type:** Content, Task pane</span></span>

## <a name="syntax"></a><span data-ttu-id="f6bd8-105">Sintaxe</span><span class="sxs-lookup"><span data-stu-id="f6bd8-105">Syntax</span></span>

```XML
<Method Name="string"/>
```

## <a name="contained-in"></a><span data-ttu-id="f6bd8-106">Contido em</span><span class="sxs-lookup"><span data-stu-id="f6bd8-106">Contained in</span></span>

[<span data-ttu-id="f6bd8-107">Methods</span><span class="sxs-lookup"><span data-stu-id="f6bd8-107">Methods</span></span>](methods.md)

## <a name="attributes"></a><span data-ttu-id="f6bd8-108">Atributos</span><span class="sxs-lookup"><span data-stu-id="f6bd8-108">Attributes</span></span>

|<span data-ttu-id="f6bd8-109">**Atributo**</span><span class="sxs-lookup"><span data-stu-id="f6bd8-109">**Attribute**</span></span>|<span data-ttu-id="f6bd8-110">**Tipo**</span><span class="sxs-lookup"><span data-stu-id="f6bd8-110">**Type**</span></span>|<span data-ttu-id="f6bd8-111">**Obrigatório**</span><span class="sxs-lookup"><span data-stu-id="f6bd8-111">**Required**</span></span>|<span data-ttu-id="f6bd8-112">**Descrição**</span><span class="sxs-lookup"><span data-stu-id="f6bd8-112">**Description**</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="f6bd8-113">Nome</span><span class="sxs-lookup"><span data-stu-id="f6bd8-113">Name</span></span>|<span data-ttu-id="f6bd8-114">cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="f6bd8-114">string</span></span>|<span data-ttu-id="f6bd8-115">obrigatório</span><span class="sxs-lookup"><span data-stu-id="f6bd8-115">required</span></span>|<span data-ttu-id="f6bd8-116">Especifica o nome do método necessário qualificado com seu objeto pai.</span><span class="sxs-lookup"><span data-stu-id="f6bd8-116">Specifies the name of the required method qualified with its parent object.</span></span> <span data-ttu-id="f6bd8-117">Por exemplo, para especificar o `getSelectedDataAsync` método, você deve especificar `"Document.getSelectedDataAsync"`.</span><span class="sxs-lookup"><span data-stu-id="f6bd8-117">For example, to specify the `getSelectedDataAsync` method, you must specify `"Document.getSelectedDataAsync"`.</span></span>|

## <a name="remarks"></a><span data-ttu-id="f6bd8-118">Comentários</span><span class="sxs-lookup"><span data-stu-id="f6bd8-118">Remarks</span></span>

<span data-ttu-id="f6bd8-119">Os `Methods` elementos `Method` e não são suportados por suplementos de email. Para obter mais informações sobre conjuntos de requisitos, confira [versões do Office e conjuntos de requisitos](../../develop/office-versions-and-requirement-sets.md).</span><span class="sxs-lookup"><span data-stu-id="f6bd8-119">The `Methods` and `Method` elements aren't supported by mail add-ins. For more information about requirement sets, see [Office versions and requirement sets](../../develop/office-versions-and-requirement-sets.md).</span></span>

> [!IMPORTANT]
> <span data-ttu-id="f6bd8-120">Como não há forma de especificar o requisito de versão mínimo de métodos individuais, para verificar se um método está disponível no tempo de execução, você também deve usar uma instrução **if** ao chamar esse método no script do suplemento.</span><span class="sxs-lookup"><span data-stu-id="f6bd8-120">Because there is no way to specify the minimum version requirement for individual methods, to make sure that a method is available at runtime, you should also use an **if** statement when calling that method in the script of your add-in.</span></span> <span data-ttu-id="f6bd8-121">Para obter mais informações sobre como fazer isso, consulte [Understanding the Office JavaScript API](../../develop/understanding-the-javascript-api-for-office.md).</span><span class="sxs-lookup"><span data-stu-id="f6bd8-121">For more information about how to do this, see [Understanding the Office JavaScript API](../../develop/understanding-the-javascript-api-for-office.md).</span></span>
