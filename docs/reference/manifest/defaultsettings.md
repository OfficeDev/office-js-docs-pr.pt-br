---
title: Elemento DefaultSettings no arquivo de manifesto
description: Especifica a localização de origem padrão e outras configurações padrão para o suplemento de conteúdo ou de painel de tarefas.
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: ace4f971d342f98d0aca5c21a7a48ceaf2563a2f
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 06/08/2020
ms.locfileid: "44611579"
---
# <a name="defaultsettings-element"></a><span data-ttu-id="c68d0-103">Elemento DefaultSettings</span><span class="sxs-lookup"><span data-stu-id="c68d0-103">DefaultSettings element</span></span>

<span data-ttu-id="c68d0-104">Especifica a localização de origem padrão e outras configurações padrão para o suplemento de conteúdo ou de painel de tarefas.</span><span class="sxs-lookup"><span data-stu-id="c68d0-104">Specifies the default source location and other default settings for your content or task pane add-in.</span></span>

<span data-ttu-id="c68d0-105">**Tipo de suplemento:** Conteúdo, Painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="c68d0-105">**Add-in type:** Content, Task pane</span></span>

## <a name="syntax"></a><span data-ttu-id="c68d0-106">Sintaxe</span><span class="sxs-lookup"><span data-stu-id="c68d0-106">Syntax</span></span>

```XML
<DefaultSettings>
  ...
</DefaultSettings>
```

## <a name="contained-in"></a><span data-ttu-id="c68d0-107">Contido em</span><span class="sxs-lookup"><span data-stu-id="c68d0-107">Contained in</span></span>

[<span data-ttu-id="c68d0-108">OfficeApp</span><span class="sxs-lookup"><span data-stu-id="c68d0-108">OfficeApp</span></span>](officeapp.md)

## <a name="can-contain"></a><span data-ttu-id="c68d0-109">Pode conter</span><span class="sxs-lookup"><span data-stu-id="c68d0-109">Can contain</span></span>

|<span data-ttu-id="c68d0-110">**Elemento**</span><span class="sxs-lookup"><span data-stu-id="c68d0-110">**Element**</span></span>|<span data-ttu-id="c68d0-111">**Content**</span><span class="sxs-lookup"><span data-stu-id="c68d0-111">**Content**</span></span>|<span data-ttu-id="c68d0-112">**Email**</span><span class="sxs-lookup"><span data-stu-id="c68d0-112">**Mail**</span></span>|<span data-ttu-id="c68d0-113">**TaskPane**</span><span class="sxs-lookup"><span data-stu-id="c68d0-113">**TaskPane**</span></span>|
|:-----|:-----|:-----|:-----|
|[<span data-ttu-id="c68d0-114">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="c68d0-114">SourceLocation</span></span>](sourcelocation.md)|<span data-ttu-id="c68d0-115">x</span><span class="sxs-lookup"><span data-stu-id="c68d0-115">x</span></span>||<span data-ttu-id="c68d0-116">x</span><span class="sxs-lookup"><span data-stu-id="c68d0-116">x</span></span>|
|[<span data-ttu-id="c68d0-117">RequestedWidth</span><span class="sxs-lookup"><span data-stu-id="c68d0-117">RequestedWidth</span></span>](requestedwidth.md)|<span data-ttu-id="c68d0-118">x</span><span class="sxs-lookup"><span data-stu-id="c68d0-118">x</span></span>|||
|[<span data-ttu-id="c68d0-119">RequestedHeight</span><span class="sxs-lookup"><span data-stu-id="c68d0-119">RequestedHeight</span></span>](requestedheight.md)|<span data-ttu-id="c68d0-120">x</span><span class="sxs-lookup"><span data-stu-id="c68d0-120">x</span></span>|||

## <a name="remarks"></a><span data-ttu-id="c68d0-121">Comentários</span><span class="sxs-lookup"><span data-stu-id="c68d0-121">Remarks</span></span>

<span data-ttu-id="c68d0-122">O local de origem e outras configurações no elemento **DefaultSettings** só se aplicam a suplementos de conteúdo e de painel de tarefas. Para suplementos de email, você especifica os locais padrão para arquivos de origem e outras configurações padrão no elemento [FormSettings](formsettings.md) .</span><span class="sxs-lookup"><span data-stu-id="c68d0-122">The source location and other settings in the **DefaultSettings** element apply only to content and task pane add-ins. For mail add-ins, you specify the default locations for source files and other default settings in the [FormSettings](formsettings.md) element.</span></span>

