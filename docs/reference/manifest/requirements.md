---
title: Elemento Requirements no arquivo de manifesto
description: O elemento requirements especifica o conjunto mínimo de requisitos e os métodos que o suplemento do Office precisa para ativar.
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: 319ddc59901c524ed1cee580a81cff749ad570db
ms.sourcegitcommit: 9609bd5b4982cdaa2ea7637709a78a45835ffb19
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 08/28/2020
ms.locfileid: "47292269"
---
# <a name="requirements-element"></a><span data-ttu-id="c88a3-103">Elemento Requirements</span><span class="sxs-lookup"><span data-stu-id="c88a3-103">Requirements element</span></span>

<span data-ttu-id="c88a3-104">Especifica o conjunto mínimo de requisitos da API JavaScript do Office ([conjuntos de requisitos](../../develop/office-versions-and-requirement-sets.md#specify-office-applications-and-requirement-sets) e/ou métodos) que o suplemento do Office precisa ativar.</span><span class="sxs-lookup"><span data-stu-id="c88a3-104">Specifies the minimum set of Office JavaScript API requirements ([requirement sets](../../develop/office-versions-and-requirement-sets.md#specify-office-applications-and-requirement-sets) and/or methods) that your Office Add-in needs to activate.</span></span>

<span data-ttu-id="c88a3-105">**Tipo de suplemento:** Conteúdo, Painel de tarefas, Email</span><span class="sxs-lookup"><span data-stu-id="c88a3-105">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="c88a3-106">Sintaxe</span><span class="sxs-lookup"><span data-stu-id="c88a3-106">Syntax</span></span>

```XML
<Requirements>
   ...
</Requirements>
```

## <a name="contained-in"></a><span data-ttu-id="c88a3-107">Contido em</span><span class="sxs-lookup"><span data-stu-id="c88a3-107">Contained in</span></span>

[<span data-ttu-id="c88a3-108">OfficeApp</span><span class="sxs-lookup"><span data-stu-id="c88a3-108">OfficeApp</span></span>](officeapp.md)

## <a name="can-contain"></a><span data-ttu-id="c88a3-109">Pode conter</span><span class="sxs-lookup"><span data-stu-id="c88a3-109">Can contain</span></span>

|<span data-ttu-id="c88a3-110">Elemento</span><span class="sxs-lookup"><span data-stu-id="c88a3-110">Element</span></span>|<span data-ttu-id="c88a3-111">Conteúdo</span><span class="sxs-lookup"><span data-stu-id="c88a3-111">Content</span></span>|<span data-ttu-id="c88a3-112">Email</span><span class="sxs-lookup"><span data-stu-id="c88a3-112">Mail</span></span>|<span data-ttu-id="c88a3-113">TaskPane</span><span class="sxs-lookup"><span data-stu-id="c88a3-113">TaskPane</span></span>|
|:-----|:-----|:-----|:-----|
|[<span data-ttu-id="c88a3-114">Sets</span><span class="sxs-lookup"><span data-stu-id="c88a3-114">Sets</span></span>](sets.md)|<span data-ttu-id="c88a3-115">x</span><span class="sxs-lookup"><span data-stu-id="c88a3-115">x</span></span>|<span data-ttu-id="c88a3-116">x</span><span class="sxs-lookup"><span data-stu-id="c88a3-116">x</span></span>|<span data-ttu-id="c88a3-117">x</span><span class="sxs-lookup"><span data-stu-id="c88a3-117">x</span></span>|
|[<span data-ttu-id="c88a3-118">Métodos</span><span class="sxs-lookup"><span data-stu-id="c88a3-118">Methods</span></span>](methods.md)|<span data-ttu-id="c88a3-119">x</span><span class="sxs-lookup"><span data-stu-id="c88a3-119">x</span></span>||<span data-ttu-id="c88a3-120">x</span><span class="sxs-lookup"><span data-stu-id="c88a3-120">x</span></span>|

## <a name="remarks"></a><span data-ttu-id="c88a3-121">Comentários</span><span class="sxs-lookup"><span data-stu-id="c88a3-121">Remarks</span></span>

<span data-ttu-id="c88a3-122">Para saber mais sobre os conjuntos de requisitos, confira [Versões do Office e conjuntos de requisitos](../../develop/office-versions-and-requirement-sets.md).</span><span class="sxs-lookup"><span data-stu-id="c88a3-122">For more information about requirement sets, see [Office versions and requirement sets](../../develop/office-versions-and-requirement-sets.md).</span></span>
