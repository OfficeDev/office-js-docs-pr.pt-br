---
title: Elemento Requirements no arquivo de manifesto
description: ''
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: 364ab7c943895e1acecedba7970e54da331a2e6f
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 04/24/2019
ms.locfileid: "32450559"
---
# <a name="requirements-element"></a><span data-ttu-id="670b8-102">Elemento Requirements</span><span class="sxs-lookup"><span data-stu-id="670b8-102">Requirements element</span></span>

<span data-ttu-id="670b8-103">Especifica o conjunto mínimo de requisitos da API JavaScript para Office ([conjuntos de requisitos](/office/dev/add-ins/develop/office-versions-and-requirement-sets#specify-office-hosts-and-requirement-sets) e/ou métodos) que o Suplemento do Office precisa ativar.</span><span class="sxs-lookup"><span data-stu-id="670b8-103">Specifies the minimum set of JavaScript API for Office requirements ([requirement sets](/office/dev/add-ins/develop/office-versions-and-requirement-sets#specify-office-hosts-and-requirement-sets) and/or methods) that your Office Add-in needs to activate.</span></span>

<span data-ttu-id="670b8-104">**Tipo de suplemento:** Conteúdo, Painel de tarefas, Email</span><span class="sxs-lookup"><span data-stu-id="670b8-104">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="670b8-105">Sintaxe</span><span class="sxs-lookup"><span data-stu-id="670b8-105">Syntax</span></span>

```XML
<Requirements>
   ...
</Requirements>
```

## <a name="contained-in"></a><span data-ttu-id="670b8-106">Contido em</span><span class="sxs-lookup"><span data-stu-id="670b8-106">Contained in</span></span>

[<span data-ttu-id="670b8-107">OfficeApp</span><span class="sxs-lookup"><span data-stu-id="670b8-107">OfficeApp</span></span>](officeapp.md)

## <a name="can-contain"></a><span data-ttu-id="670b8-108">Pode conter</span><span class="sxs-lookup"><span data-stu-id="670b8-108">Can contain</span></span>

|<span data-ttu-id="670b8-109">**Elemento**</span><span class="sxs-lookup"><span data-stu-id="670b8-109">**Element**</span></span>|<span data-ttu-id="670b8-110">**Content**</span><span class="sxs-lookup"><span data-stu-id="670b8-110">**Content**</span></span>|<span data-ttu-id="670b8-111">**Email**</span><span class="sxs-lookup"><span data-stu-id="670b8-111">**Mail**</span></span>|<span data-ttu-id="670b8-112">**TaskPane**</span><span class="sxs-lookup"><span data-stu-id="670b8-112">**TaskPane**</span></span>|
|:-----|:-----|:-----|:-----|
|[<span data-ttu-id="670b8-113">Sets</span><span class="sxs-lookup"><span data-stu-id="670b8-113">Sets</span></span>](sets.md)|<span data-ttu-id="670b8-114">x</span><span class="sxs-lookup"><span data-stu-id="670b8-114">x</span></span>|<span data-ttu-id="670b8-115">x</span><span class="sxs-lookup"><span data-stu-id="670b8-115">x</span></span>|<span data-ttu-id="670b8-116">x</span><span class="sxs-lookup"><span data-stu-id="670b8-116">x</span></span>|
|[<span data-ttu-id="670b8-117">Métodos</span><span class="sxs-lookup"><span data-stu-id="670b8-117">Methods</span></span>](methods.md)|<span data-ttu-id="670b8-118">x</span><span class="sxs-lookup"><span data-stu-id="670b8-118">x</span></span>||<span data-ttu-id="670b8-119">x</span><span class="sxs-lookup"><span data-stu-id="670b8-119">x</span></span>|

## <a name="remarks"></a><span data-ttu-id="670b8-120">Comentários</span><span class="sxs-lookup"><span data-stu-id="670b8-120">Remarks</span></span>

<span data-ttu-id="670b8-121">Para saber mais sobre os conjuntos de requisitos, confira [Versões do Office e conjuntos de requisitos](/office/dev/add-ins/develop/office-versions-and-requirement-sets).</span><span class="sxs-lookup"><span data-stu-id="670b8-121">For more information about requirement sets, see [Office versions and requirement sets](/office/dev/add-ins/develop/office-versions-and-requirement-sets).</span></span>

