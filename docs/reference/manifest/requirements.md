---
title: Elemento Requirements no arquivo de manifesto
description: ''
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: 43c66118b9129c4c8ae395254ea82ef1cbcbaab1
ms.sourcegitcommit: 4079903c3cc45b7d8c041509a44e9fc38da399b1
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/11/2020
ms.locfileid: "42596456"
---
# <a name="requirements-element"></a><span data-ttu-id="fb22a-102">Elemento Requirements</span><span class="sxs-lookup"><span data-stu-id="fb22a-102">Requirements element</span></span>

<span data-ttu-id="fb22a-103">Especifica o conjunto mínimo de requisitos da API JavaScript do Office ([conjuntos de requisitos](../../develop/office-versions-and-requirement-sets.md#specify-office-hosts-and-requirement-sets) e/ou métodos) que o suplemento do Office precisa ativar.</span><span class="sxs-lookup"><span data-stu-id="fb22a-103">Specifies the minimum set of Office JavaScript API requirements ([requirement sets](../../develop/office-versions-and-requirement-sets.md#specify-office-hosts-and-requirement-sets) and/or methods) that your Office Add-in needs to activate.</span></span>

<span data-ttu-id="fb22a-104">**Tipo de suplemento:** Conteúdo, Painel de tarefas, Email</span><span class="sxs-lookup"><span data-stu-id="fb22a-104">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="fb22a-105">Sintaxe</span><span class="sxs-lookup"><span data-stu-id="fb22a-105">Syntax</span></span>

```XML
<Requirements>
   ...
</Requirements>
```

## <a name="contained-in"></a><span data-ttu-id="fb22a-106">Contido em</span><span class="sxs-lookup"><span data-stu-id="fb22a-106">Contained in</span></span>

[<span data-ttu-id="fb22a-107">OfficeApp</span><span class="sxs-lookup"><span data-stu-id="fb22a-107">OfficeApp</span></span>](officeapp.md)

## <a name="can-contain"></a><span data-ttu-id="fb22a-108">Pode conter</span><span class="sxs-lookup"><span data-stu-id="fb22a-108">Can contain</span></span>

|<span data-ttu-id="fb22a-109">**Elemento**</span><span class="sxs-lookup"><span data-stu-id="fb22a-109">**Element**</span></span>|<span data-ttu-id="fb22a-110">**Content**</span><span class="sxs-lookup"><span data-stu-id="fb22a-110">**Content**</span></span>|<span data-ttu-id="fb22a-111">**Email**</span><span class="sxs-lookup"><span data-stu-id="fb22a-111">**Mail**</span></span>|<span data-ttu-id="fb22a-112">**TaskPane**</span><span class="sxs-lookup"><span data-stu-id="fb22a-112">**TaskPane**</span></span>|
|:-----|:-----|:-----|:-----|
|[<span data-ttu-id="fb22a-113">Sets</span><span class="sxs-lookup"><span data-stu-id="fb22a-113">Sets</span></span>](sets.md)|<span data-ttu-id="fb22a-114">x</span><span class="sxs-lookup"><span data-stu-id="fb22a-114">x</span></span>|<span data-ttu-id="fb22a-115">x</span><span class="sxs-lookup"><span data-stu-id="fb22a-115">x</span></span>|<span data-ttu-id="fb22a-116">x</span><span class="sxs-lookup"><span data-stu-id="fb22a-116">x</span></span>|
|[<span data-ttu-id="fb22a-117">Métodos</span><span class="sxs-lookup"><span data-stu-id="fb22a-117">Methods</span></span>](methods.md)|<span data-ttu-id="fb22a-118">x</span><span class="sxs-lookup"><span data-stu-id="fb22a-118">x</span></span>||<span data-ttu-id="fb22a-119">x</span><span class="sxs-lookup"><span data-stu-id="fb22a-119">x</span></span>|

## <a name="remarks"></a><span data-ttu-id="fb22a-120">Comentários</span><span class="sxs-lookup"><span data-stu-id="fb22a-120">Remarks</span></span>

<span data-ttu-id="fb22a-121">Para saber mais sobre os conjuntos de requisitos, confira [Versões do Office e conjuntos de requisitos](../../develop/office-versions-and-requirement-sets.md).</span><span class="sxs-lookup"><span data-stu-id="fb22a-121">For more information about requirement sets, see [Office versions and requirement sets](../../develop/office-versions-and-requirement-sets.md).</span></span>
