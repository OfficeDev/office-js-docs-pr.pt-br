---
title: Elemento Requirements no arquivo de manifesto
description: ''
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: 3c4cb81ebd6a38ea311e8fcacfa6d5fcd3b26f68
ms.sourcegitcommit: 5d29801180f6939ec10efb778d2311be67d8b9f1
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 02/27/2020
ms.locfileid: "42325245"
---
# <a name="requirements-element"></a><span data-ttu-id="8b36d-102">Elemento Requirements</span><span class="sxs-lookup"><span data-stu-id="8b36d-102">Requirements element</span></span>

<span data-ttu-id="8b36d-103">Especifica o conjunto mínimo de requisitos da API JavaScript do Office ([conjuntos de requisitos](/office/dev/add-ins/develop/office-versions-and-requirement-sets#specify-office-hosts-and-requirement-sets) e/ou métodos) que o suplemento do Office precisa ativar.</span><span class="sxs-lookup"><span data-stu-id="8b36d-103">Specifies the minimum set of Office JavaScript API requirements ([requirement sets](/office/dev/add-ins/develop/office-versions-and-requirement-sets#specify-office-hosts-and-requirement-sets) and/or methods) that your Office Add-in needs to activate.</span></span>

<span data-ttu-id="8b36d-104">**Tipo de suplemento:** Conteúdo, Painel de tarefas, Email</span><span class="sxs-lookup"><span data-stu-id="8b36d-104">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="8b36d-105">Sintaxe</span><span class="sxs-lookup"><span data-stu-id="8b36d-105">Syntax</span></span>

```XML
<Requirements>
   ...
</Requirements>
```

## <a name="contained-in"></a><span data-ttu-id="8b36d-106">Contido em</span><span class="sxs-lookup"><span data-stu-id="8b36d-106">Contained in</span></span>

[<span data-ttu-id="8b36d-107">OfficeApp</span><span class="sxs-lookup"><span data-stu-id="8b36d-107">OfficeApp</span></span>](officeapp.md)

## <a name="can-contain"></a><span data-ttu-id="8b36d-108">Pode conter</span><span class="sxs-lookup"><span data-stu-id="8b36d-108">Can contain</span></span>

|<span data-ttu-id="8b36d-109">**Elemento**</span><span class="sxs-lookup"><span data-stu-id="8b36d-109">**Element**</span></span>|<span data-ttu-id="8b36d-110">**Content**</span><span class="sxs-lookup"><span data-stu-id="8b36d-110">**Content**</span></span>|<span data-ttu-id="8b36d-111">**Email**</span><span class="sxs-lookup"><span data-stu-id="8b36d-111">**Mail**</span></span>|<span data-ttu-id="8b36d-112">**TaskPane**</span><span class="sxs-lookup"><span data-stu-id="8b36d-112">**TaskPane**</span></span>|
|:-----|:-----|:-----|:-----|
|[<span data-ttu-id="8b36d-113">Sets</span><span class="sxs-lookup"><span data-stu-id="8b36d-113">Sets</span></span>](sets.md)|<span data-ttu-id="8b36d-114">x</span><span class="sxs-lookup"><span data-stu-id="8b36d-114">x</span></span>|<span data-ttu-id="8b36d-115">x</span><span class="sxs-lookup"><span data-stu-id="8b36d-115">x</span></span>|<span data-ttu-id="8b36d-116">x</span><span class="sxs-lookup"><span data-stu-id="8b36d-116">x</span></span>|
|[<span data-ttu-id="8b36d-117">Métodos</span><span class="sxs-lookup"><span data-stu-id="8b36d-117">Methods</span></span>](methods.md)|<span data-ttu-id="8b36d-118">x</span><span class="sxs-lookup"><span data-stu-id="8b36d-118">x</span></span>||<span data-ttu-id="8b36d-119">x</span><span class="sxs-lookup"><span data-stu-id="8b36d-119">x</span></span>|

## <a name="remarks"></a><span data-ttu-id="8b36d-120">Comentários</span><span class="sxs-lookup"><span data-stu-id="8b36d-120">Remarks</span></span>

<span data-ttu-id="8b36d-121">Para saber mais sobre os conjuntos de requisitos, confira [Versões do Office e conjuntos de requisitos](/office/dev/add-ins/develop/office-versions-and-requirement-sets).</span><span class="sxs-lookup"><span data-stu-id="8b36d-121">For more information about requirement sets, see [Office versions and requirement sets](/office/dev/add-ins/develop/office-versions-and-requirement-sets).</span></span>

