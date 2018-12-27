---
title: Elemento Requirements no arquivo de manifesto
description: ''
ms.date: 10/09/2018
ms.openlocfilehash: 2544e9b01b2d4d3ddc0a0c6238b4a5b0e6c4f832
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 12/22/2018
ms.locfileid: "27432701"
---
# <a name="requirements-element"></a><span data-ttu-id="cab82-102">Elemento Requirements</span><span class="sxs-lookup"><span data-stu-id="cab82-102">Requirements element</span></span>

<span data-ttu-id="cab82-103">Especifica o conjunto mínimo de requisitos da API JavaScript para Office ([conjuntos de requisitos](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets#specify-office-hosts-and-requirement-sets) e/ou métodos) que o Suplemento do Office precisa ativar.</span><span class="sxs-lookup"><span data-stu-id="cab82-103">Specifies the minimum set of JavaScript API for Office requirements ([requirement sets](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets#specify-office-hosts-and-requirement-sets) and/or methods) that your Office Add-in needs to activate.</span></span>

<span data-ttu-id="cab82-104">**Tipo de suplemento:** Conteúdo, Painel de tarefas, Email</span><span class="sxs-lookup"><span data-stu-id="cab82-104">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="cab82-105">Sintaxe</span><span class="sxs-lookup"><span data-stu-id="cab82-105">Syntax</span></span>

```XML
<Requirements>
   ...
</Requirements>
```

## <a name="contained-in"></a><span data-ttu-id="cab82-106">Contido em</span><span class="sxs-lookup"><span data-stu-id="cab82-106">Contained in</span></span>

[<span data-ttu-id="cab82-107">OfficeApp</span><span class="sxs-lookup"><span data-stu-id="cab82-107">OfficeApp</span></span>](officeapp.md)

## <a name="can-contain"></a><span data-ttu-id="cab82-108">Pode conter</span><span class="sxs-lookup"><span data-stu-id="cab82-108">Can contain</span></span>

|<span data-ttu-id="cab82-109">**Element**</span><span class="sxs-lookup"><span data-stu-id="cab82-109">**Element**</span></span>|<span data-ttu-id="cab82-110">**Conteúdo**</span><span class="sxs-lookup"><span data-stu-id="cab82-110">**Content**</span></span>|<span data-ttu-id="cab82-111">**Email**</span><span class="sxs-lookup"><span data-stu-id="cab82-111">**Mail**</span></span>|<span data-ttu-id="cab82-112">**TaskPane**</span><span class="sxs-lookup"><span data-stu-id="cab82-112">**TaskPane**</span></span>|
|:-----|:-----|:-----|:-----|
|[<span data-ttu-id="cab82-113">Sets</span><span class="sxs-lookup"><span data-stu-id="cab82-113">Sets</span></span>](sets.md)|<span data-ttu-id="cab82-114">x</span><span class="sxs-lookup"><span data-stu-id="cab82-114">x</span></span>|<span data-ttu-id="cab82-115">x</span><span class="sxs-lookup"><span data-stu-id="cab82-115">x</span></span>|<span data-ttu-id="cab82-116">x</span><span class="sxs-lookup"><span data-stu-id="cab82-116">x</span></span>|
|[<span data-ttu-id="cab82-117">Métodos</span><span class="sxs-lookup"><span data-stu-id="cab82-117">Methods</span></span>](methods.md)|<span data-ttu-id="cab82-118">x</span><span class="sxs-lookup"><span data-stu-id="cab82-118">x</span></span>||<span data-ttu-id="cab82-119">x</span><span class="sxs-lookup"><span data-stu-id="cab82-119">x</span></span>|

## <a name="remarks"></a><span data-ttu-id="cab82-120">Comentários</span><span class="sxs-lookup"><span data-stu-id="cab82-120">Remarks</span></span>

<span data-ttu-id="cab82-121">Para saber mais sobre os conjuntos de requisitos, confira [Versões do Office e conjuntos de requisitos](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets).</span><span class="sxs-lookup"><span data-stu-id="cab82-121">For more information about available requirement sets, see [Office add-in requirement sets](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets).</span></span>

