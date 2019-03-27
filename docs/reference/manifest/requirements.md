---
title: Elemento Requirements no arquivo de manifesto
description: ''
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: 364ab7c943895e1acecedba7970e54da331a2e6f
ms.sourcegitcommit: a2950492a2337de3180b713f5693fe82dbdd6a17
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/27/2019
ms.locfileid: "30870363"
---
# <a name="requirements-element"></a><span data-ttu-id="20730-102">Elemento Requirements</span><span class="sxs-lookup"><span data-stu-id="20730-102">Requirements element</span></span>

<span data-ttu-id="20730-103">Especifica o conjunto mínimo de requisitos da API JavaScript para Office ([conjuntos de requisitos](/office/dev/add-ins/develop/office-versions-and-requirement-sets#specify-office-hosts-and-requirement-sets) e/ou métodos) que o Suplemento do Office precisa ativar.</span><span class="sxs-lookup"><span data-stu-id="20730-103">Specifies the minimum set of JavaScript API for Office requirements ([requirement sets](/office/dev/add-ins/develop/office-versions-and-requirement-sets#specify-office-hosts-and-requirement-sets) and/or methods) that your Office Add-in needs to activate.</span></span>

<span data-ttu-id="20730-104">**Tipo de suplemento:** Conteúdo, Painel de tarefas, Email</span><span class="sxs-lookup"><span data-stu-id="20730-104">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="20730-105">Sintaxe</span><span class="sxs-lookup"><span data-stu-id="20730-105">Syntax</span></span>

```XML
<Requirements>
   ...
</Requirements>
```

## <a name="contained-in"></a><span data-ttu-id="20730-106">Contido em</span><span class="sxs-lookup"><span data-stu-id="20730-106">Contained in</span></span>

[<span data-ttu-id="20730-107">OfficeApp</span><span class="sxs-lookup"><span data-stu-id="20730-107">OfficeApp</span></span>](officeapp.md)

## <a name="can-contain"></a><span data-ttu-id="20730-108">Pode conter</span><span class="sxs-lookup"><span data-stu-id="20730-108">Can contain</span></span>

|<span data-ttu-id="20730-109">**Elemento**</span><span class="sxs-lookup"><span data-stu-id="20730-109">**Element**</span></span>|<span data-ttu-id="20730-110">**Content**</span><span class="sxs-lookup"><span data-stu-id="20730-110">**Content**</span></span>|<span data-ttu-id="20730-111">**Email**</span><span class="sxs-lookup"><span data-stu-id="20730-111">**Mail**</span></span>|<span data-ttu-id="20730-112">**TaskPane**</span><span class="sxs-lookup"><span data-stu-id="20730-112">**TaskPane**</span></span>|
|:-----|:-----|:-----|:-----|
|[<span data-ttu-id="20730-113">Sets</span><span class="sxs-lookup"><span data-stu-id="20730-113">Sets</span></span>](sets.md)|<span data-ttu-id="20730-114">x</span><span class="sxs-lookup"><span data-stu-id="20730-114">x</span></span>|<span data-ttu-id="20730-115">x</span><span class="sxs-lookup"><span data-stu-id="20730-115">x</span></span>|<span data-ttu-id="20730-116">x</span><span class="sxs-lookup"><span data-stu-id="20730-116">x</span></span>|
|[<span data-ttu-id="20730-117">Métodos</span><span class="sxs-lookup"><span data-stu-id="20730-117">Methods</span></span>](methods.md)|<span data-ttu-id="20730-118">x</span><span class="sxs-lookup"><span data-stu-id="20730-118">x</span></span>||<span data-ttu-id="20730-119">x</span><span class="sxs-lookup"><span data-stu-id="20730-119">x</span></span>|

## <a name="remarks"></a><span data-ttu-id="20730-120">Comentários</span><span class="sxs-lookup"><span data-stu-id="20730-120">Remarks</span></span>

<span data-ttu-id="20730-121">Para saber mais sobre os conjuntos de requisitos, confira [Versões do Office e conjuntos de requisitos](/office/dev/add-ins/develop/office-versions-and-requirement-sets).</span><span class="sxs-lookup"><span data-stu-id="20730-121">For more information about requirement sets, see [Office versions and requirement sets](/office/dev/add-ins/develop/office-versions-and-requirement-sets).</span></span>

