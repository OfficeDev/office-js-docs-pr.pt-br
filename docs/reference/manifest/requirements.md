---
title: Elemento Requirements no arquivo de manifesto
description: O elemento requirements especifica o conjunto mínimo de requisitos e os métodos que o suplemento do Office precisa para ativar.
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: a3f41a763ec820a6c766e6a32b26e55ad34996f7
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/17/2020
ms.locfileid: "42720446"
---
# <a name="requirements-element"></a><span data-ttu-id="c84e0-103">Elemento Requirements</span><span class="sxs-lookup"><span data-stu-id="c84e0-103">Requirements element</span></span>

<span data-ttu-id="c84e0-104">Especifica o conjunto mínimo de requisitos da API JavaScript do Office ([conjuntos de requisitos](../../develop/office-versions-and-requirement-sets.md#specify-office-hosts-and-requirement-sets) e/ou métodos) que o suplemento do Office precisa ativar.</span><span class="sxs-lookup"><span data-stu-id="c84e0-104">Specifies the minimum set of Office JavaScript API requirements ([requirement sets](../../develop/office-versions-and-requirement-sets.md#specify-office-hosts-and-requirement-sets) and/or methods) that your Office Add-in needs to activate.</span></span>

<span data-ttu-id="c84e0-105">**Tipo de suplemento:** Conteúdo, Painel de tarefas, Email</span><span class="sxs-lookup"><span data-stu-id="c84e0-105">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="c84e0-106">Sintaxe</span><span class="sxs-lookup"><span data-stu-id="c84e0-106">Syntax</span></span>

```XML
<Requirements>
   ...
</Requirements>
```

## <a name="contained-in"></a><span data-ttu-id="c84e0-107">Contido em</span><span class="sxs-lookup"><span data-stu-id="c84e0-107">Contained in</span></span>

[<span data-ttu-id="c84e0-108">OfficeApp</span><span class="sxs-lookup"><span data-stu-id="c84e0-108">OfficeApp</span></span>](officeapp.md)

## <a name="can-contain"></a><span data-ttu-id="c84e0-109">Pode conter</span><span class="sxs-lookup"><span data-stu-id="c84e0-109">Can contain</span></span>

|<span data-ttu-id="c84e0-110">**Elemento**</span><span class="sxs-lookup"><span data-stu-id="c84e0-110">**Element**</span></span>|<span data-ttu-id="c84e0-111">**Content**</span><span class="sxs-lookup"><span data-stu-id="c84e0-111">**Content**</span></span>|<span data-ttu-id="c84e0-112">**Email**</span><span class="sxs-lookup"><span data-stu-id="c84e0-112">**Mail**</span></span>|<span data-ttu-id="c84e0-113">**TaskPane**</span><span class="sxs-lookup"><span data-stu-id="c84e0-113">**TaskPane**</span></span>|
|:-----|:-----|:-----|:-----|
|[<span data-ttu-id="c84e0-114">Sets</span><span class="sxs-lookup"><span data-stu-id="c84e0-114">Sets</span></span>](sets.md)|<span data-ttu-id="c84e0-115">x</span><span class="sxs-lookup"><span data-stu-id="c84e0-115">x</span></span>|<span data-ttu-id="c84e0-116">x</span><span class="sxs-lookup"><span data-stu-id="c84e0-116">x</span></span>|<span data-ttu-id="c84e0-117">x</span><span class="sxs-lookup"><span data-stu-id="c84e0-117">x</span></span>|
|[<span data-ttu-id="c84e0-118">Métodos</span><span class="sxs-lookup"><span data-stu-id="c84e0-118">Methods</span></span>](methods.md)|<span data-ttu-id="c84e0-119">x</span><span class="sxs-lookup"><span data-stu-id="c84e0-119">x</span></span>||<span data-ttu-id="c84e0-120">x</span><span class="sxs-lookup"><span data-stu-id="c84e0-120">x</span></span>|

## <a name="remarks"></a><span data-ttu-id="c84e0-121">Comentários</span><span class="sxs-lookup"><span data-stu-id="c84e0-121">Remarks</span></span>

<span data-ttu-id="c84e0-122">Para saber mais sobre os conjuntos de requisitos, confira [Versões do Office e conjuntos de requisitos](../../develop/office-versions-and-requirement-sets.md).</span><span class="sxs-lookup"><span data-stu-id="c84e0-122">For more information about requirement sets, see [Office versions and requirement sets](../../develop/office-versions-and-requirement-sets.md).</span></span>
