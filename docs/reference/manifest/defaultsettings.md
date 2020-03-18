---
title: Elemento DefaultSettings no arquivo de manifesto
description: Especifica a localização de origem padrão e outras configurações padrão para o suplemento de conteúdo ou de painel de tarefas.
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: b97f692a1fd39e4b1f55080f6ed77e623be0000c
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/17/2020
ms.locfileid: "42718367"
---
# <a name="defaultsettings-element"></a><span data-ttu-id="eae7d-103">Elemento DefaultSettings</span><span class="sxs-lookup"><span data-stu-id="eae7d-103">DefaultSettings element</span></span>

<span data-ttu-id="eae7d-104">Especifica a localização de origem padrão e outras configurações padrão para o suplemento de conteúdo ou de painel de tarefas.</span><span class="sxs-lookup"><span data-stu-id="eae7d-104">Specifies the default source location and other default settings for your content or task pane add-in.</span></span>

<span data-ttu-id="eae7d-105">**Tipo de suplemento:** Conteúdo, Painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="eae7d-105">**Add-in type:** Content, Task pane</span></span>

## <a name="syntax"></a><span data-ttu-id="eae7d-106">Sintaxe</span><span class="sxs-lookup"><span data-stu-id="eae7d-106">Syntax</span></span>

```XML
<DefaultSettings>
  ...
</DefaultSettings>
```

## <a name="contained-in"></a><span data-ttu-id="eae7d-107">Contido em</span><span class="sxs-lookup"><span data-stu-id="eae7d-107">Contained in</span></span>

[<span data-ttu-id="eae7d-108">OfficeApp</span><span class="sxs-lookup"><span data-stu-id="eae7d-108">OfficeApp</span></span>](officeapp.md)

## <a name="can-contain"></a><span data-ttu-id="eae7d-109">Pode conter</span><span class="sxs-lookup"><span data-stu-id="eae7d-109">Can contain</span></span>

|<span data-ttu-id="eae7d-110">**Elemento**</span><span class="sxs-lookup"><span data-stu-id="eae7d-110">**Element**</span></span>|<span data-ttu-id="eae7d-111">**Content**</span><span class="sxs-lookup"><span data-stu-id="eae7d-111">**Content**</span></span>|<span data-ttu-id="eae7d-112">**Email**</span><span class="sxs-lookup"><span data-stu-id="eae7d-112">**Mail**</span></span>|<span data-ttu-id="eae7d-113">**TaskPane**</span><span class="sxs-lookup"><span data-stu-id="eae7d-113">**TaskPane**</span></span>|
|:-----|:-----|:-----|:-----|
|[<span data-ttu-id="eae7d-114">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="eae7d-114">SourceLocation</span></span>](sourcelocation.md)|<span data-ttu-id="eae7d-115">x</span><span class="sxs-lookup"><span data-stu-id="eae7d-115">x</span></span>||<span data-ttu-id="eae7d-116">x</span><span class="sxs-lookup"><span data-stu-id="eae7d-116">x</span></span>|
|[<span data-ttu-id="eae7d-117">RequestedWidth</span><span class="sxs-lookup"><span data-stu-id="eae7d-117">RequestedWidth</span></span>](requestedwidth.md)|<span data-ttu-id="eae7d-118">x</span><span class="sxs-lookup"><span data-stu-id="eae7d-118">x</span></span>|||
|[<span data-ttu-id="eae7d-119">RequestedHeight</span><span class="sxs-lookup"><span data-stu-id="eae7d-119">RequestedHeight</span></span>](requestedheight.md)|<span data-ttu-id="eae7d-120">x</span><span class="sxs-lookup"><span data-stu-id="eae7d-120">x</span></span>|||

## <a name="remarks"></a><span data-ttu-id="eae7d-121">Comentários</span><span class="sxs-lookup"><span data-stu-id="eae7d-121">Remarks</span></span>

<span data-ttu-id="eae7d-122">O local de origem e outras configurações no elemento **DefaultSettings** só se aplicam a suplementos de conteúdo e de painel de tarefas. Para suplementos de email, você especifica os locais padrão para arquivos de origem e outras configurações padrão no elemento [FormSettings](formsettings.md) .</span><span class="sxs-lookup"><span data-stu-id="eae7d-122">The source location and other settings in the **DefaultSettings** element apply only to content and task pane add-ins. For mail add-ins, you specify the default locations for source files and other default settings in the [FormSettings](formsettings.md) element.</span></span>

