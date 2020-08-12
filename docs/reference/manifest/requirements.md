---
title: Elemento Requirements no arquivo de manifesto
description: O elemento requirements especifica o conjunto mínimo de requisitos e os métodos que o suplemento do Office precisa para ativar.
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: c6a9a7b5923401fc2551f239b2c6cbc0d1e90755
ms.sourcegitcommit: cc6886b47c84ac37a3c957ff85dd0ed526ca5e43
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 08/12/2020
ms.locfileid: "46641316"
---
# <a name="requirements-element"></a><span data-ttu-id="50aa1-103">Elemento Requirements</span><span class="sxs-lookup"><span data-stu-id="50aa1-103">Requirements element</span></span>

<span data-ttu-id="50aa1-104">Especifica o conjunto mínimo de requisitos da API JavaScript do Office ([conjuntos de requisitos](../../develop/office-versions-and-requirement-sets.md#specify-office-hosts-and-requirement-sets) e/ou métodos) que o suplemento do Office precisa ativar.</span><span class="sxs-lookup"><span data-stu-id="50aa1-104">Specifies the minimum set of Office JavaScript API requirements ([requirement sets](../../develop/office-versions-and-requirement-sets.md#specify-office-hosts-and-requirement-sets) and/or methods) that your Office Add-in needs to activate.</span></span>

<span data-ttu-id="50aa1-105">**Tipo de suplemento:** Conteúdo, Painel de tarefas, Email</span><span class="sxs-lookup"><span data-stu-id="50aa1-105">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="50aa1-106">Sintaxe</span><span class="sxs-lookup"><span data-stu-id="50aa1-106">Syntax</span></span>

```XML
<Requirements>
   ...
</Requirements>
```

## <a name="contained-in"></a><span data-ttu-id="50aa1-107">Contido em</span><span class="sxs-lookup"><span data-stu-id="50aa1-107">Contained in</span></span>

[<span data-ttu-id="50aa1-108">OfficeApp</span><span class="sxs-lookup"><span data-stu-id="50aa1-108">OfficeApp</span></span>](officeapp.md)

## <a name="can-contain"></a><span data-ttu-id="50aa1-109">Pode conter</span><span class="sxs-lookup"><span data-stu-id="50aa1-109">Can contain</span></span>

|<span data-ttu-id="50aa1-110">Elemento</span><span class="sxs-lookup"><span data-stu-id="50aa1-110">Element</span></span>|<span data-ttu-id="50aa1-111">Conteúdo</span><span class="sxs-lookup"><span data-stu-id="50aa1-111">Content</span></span>|<span data-ttu-id="50aa1-112">Email</span><span class="sxs-lookup"><span data-stu-id="50aa1-112">Mail</span></span>|<span data-ttu-id="50aa1-113">TaskPane</span><span class="sxs-lookup"><span data-stu-id="50aa1-113">TaskPane</span></span>|
|:-----|:-----|:-----|:-----|
|[<span data-ttu-id="50aa1-114">Sets</span><span class="sxs-lookup"><span data-stu-id="50aa1-114">Sets</span></span>](sets.md)|<span data-ttu-id="50aa1-115">x</span><span class="sxs-lookup"><span data-stu-id="50aa1-115">x</span></span>|<span data-ttu-id="50aa1-116">x</span><span class="sxs-lookup"><span data-stu-id="50aa1-116">x</span></span>|<span data-ttu-id="50aa1-117">x</span><span class="sxs-lookup"><span data-stu-id="50aa1-117">x</span></span>|
|[<span data-ttu-id="50aa1-118">Métodos</span><span class="sxs-lookup"><span data-stu-id="50aa1-118">Methods</span></span>](methods.md)|<span data-ttu-id="50aa1-119">x</span><span class="sxs-lookup"><span data-stu-id="50aa1-119">x</span></span>||<span data-ttu-id="50aa1-120">x</span><span class="sxs-lookup"><span data-stu-id="50aa1-120">x</span></span>|

## <a name="remarks"></a><span data-ttu-id="50aa1-121">Comentários</span><span class="sxs-lookup"><span data-stu-id="50aa1-121">Remarks</span></span>

<span data-ttu-id="50aa1-122">Para saber mais sobre os conjuntos de requisitos, confira [Versões do Office e conjuntos de requisitos](../../develop/office-versions-and-requirement-sets.md).</span><span class="sxs-lookup"><span data-stu-id="50aa1-122">For more information about requirement sets, see [Office versions and requirement sets](../../develop/office-versions-and-requirement-sets.md).</span></span>
