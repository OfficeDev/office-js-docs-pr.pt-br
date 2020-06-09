---
title: Elemento Requirements no arquivo de manifesto
description: O elemento requirements especifica o conjunto mínimo de requisitos e os métodos que o suplemento do Office precisa para ativar.
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: 586f05ec68257462cb64a96abf2a34eb31861a5c
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 06/08/2020
ms.locfileid: "44611712"
---
# <a name="requirements-element"></a><span data-ttu-id="04052-103">Elemento Requirements</span><span class="sxs-lookup"><span data-stu-id="04052-103">Requirements element</span></span>

<span data-ttu-id="04052-104">Especifica o conjunto mínimo de requisitos da API JavaScript do Office ([conjuntos de requisitos](../../develop/office-versions-and-requirement-sets.md#specify-office-hosts-and-requirement-sets) e/ou métodos) que o suplemento do Office precisa ativar.</span><span class="sxs-lookup"><span data-stu-id="04052-104">Specifies the minimum set of Office JavaScript API requirements ([requirement sets](../../develop/office-versions-and-requirement-sets.md#specify-office-hosts-and-requirement-sets) and/or methods) that your Office Add-in needs to activate.</span></span>

<span data-ttu-id="04052-105">**Tipo de suplemento:** Conteúdo, Painel de tarefas, Email</span><span class="sxs-lookup"><span data-stu-id="04052-105">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="04052-106">Sintaxe</span><span class="sxs-lookup"><span data-stu-id="04052-106">Syntax</span></span>

```XML
<Requirements>
   ...
</Requirements>
```

## <a name="contained-in"></a><span data-ttu-id="04052-107">Contido em</span><span class="sxs-lookup"><span data-stu-id="04052-107">Contained in</span></span>

[<span data-ttu-id="04052-108">OfficeApp</span><span class="sxs-lookup"><span data-stu-id="04052-108">OfficeApp</span></span>](officeapp.md)

## <a name="can-contain"></a><span data-ttu-id="04052-109">Pode conter</span><span class="sxs-lookup"><span data-stu-id="04052-109">Can contain</span></span>

|<span data-ttu-id="04052-110">**Elemento**</span><span class="sxs-lookup"><span data-stu-id="04052-110">**Element**</span></span>|<span data-ttu-id="04052-111">**Content**</span><span class="sxs-lookup"><span data-stu-id="04052-111">**Content**</span></span>|<span data-ttu-id="04052-112">**Email**</span><span class="sxs-lookup"><span data-stu-id="04052-112">**Mail**</span></span>|<span data-ttu-id="04052-113">**TaskPane**</span><span class="sxs-lookup"><span data-stu-id="04052-113">**TaskPane**</span></span>|
|:-----|:-----|:-----|:-----|
|[<span data-ttu-id="04052-114">Sets</span><span class="sxs-lookup"><span data-stu-id="04052-114">Sets</span></span>](sets.md)|<span data-ttu-id="04052-115">x</span><span class="sxs-lookup"><span data-stu-id="04052-115">x</span></span>|<span data-ttu-id="04052-116">x</span><span class="sxs-lookup"><span data-stu-id="04052-116">x</span></span>|<span data-ttu-id="04052-117">x</span><span class="sxs-lookup"><span data-stu-id="04052-117">x</span></span>|
|[<span data-ttu-id="04052-118">Métodos</span><span class="sxs-lookup"><span data-stu-id="04052-118">Methods</span></span>](methods.md)|<span data-ttu-id="04052-119">x</span><span class="sxs-lookup"><span data-stu-id="04052-119">x</span></span>||<span data-ttu-id="04052-120">x</span><span class="sxs-lookup"><span data-stu-id="04052-120">x</span></span>|

## <a name="remarks"></a><span data-ttu-id="04052-121">Comentários</span><span class="sxs-lookup"><span data-stu-id="04052-121">Remarks</span></span>

<span data-ttu-id="04052-122">Para saber mais sobre os conjuntos de requisitos, confira [Versões do Office e conjuntos de requisitos](../../develop/office-versions-and-requirement-sets.md).</span><span class="sxs-lookup"><span data-stu-id="04052-122">For more information about requirement sets, see [Office versions and requirement sets](../../develop/office-versions-and-requirement-sets.md).</span></span>
