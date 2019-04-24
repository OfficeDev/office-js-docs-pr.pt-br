---
title: Elemento DefaultSettings no arquivo de manifesto
description: ''
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: 199acf8be888ba51fda83d159937a74685ca48e0
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 04/24/2019
ms.locfileid: "32450622"
---
# <a name="defaultsettings-element"></a><span data-ttu-id="5ca6f-102">Elemento DefaultSettings</span><span class="sxs-lookup"><span data-stu-id="5ca6f-102">DefaultSettings element</span></span>

<span data-ttu-id="5ca6f-103">Especifica a localização de origem padrão e outras configurações padrão para o suplemento de conteúdo ou de painel de tarefas.</span><span class="sxs-lookup"><span data-stu-id="5ca6f-103">Specifies the default source location and other default settings for your content or task pane add-in.</span></span>

<span data-ttu-id="5ca6f-104">**Tipo de suplemento:** Conteúdo, Painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="5ca6f-104">**Add-in type:** Content, Task pane</span></span>

## <a name="syntax"></a><span data-ttu-id="5ca6f-105">Sintaxe</span><span class="sxs-lookup"><span data-stu-id="5ca6f-105">Syntax</span></span>

```XML
<DefaultSettings>
  ...
</DefaultSettings>
```

## <a name="contained-in"></a><span data-ttu-id="5ca6f-106">Contido em</span><span class="sxs-lookup"><span data-stu-id="5ca6f-106">Contained in</span></span>

[<span data-ttu-id="5ca6f-107">OfficeApp</span><span class="sxs-lookup"><span data-stu-id="5ca6f-107">OfficeApp</span></span>](officeapp.md)

## <a name="can-contain"></a><span data-ttu-id="5ca6f-108">Pode conter</span><span class="sxs-lookup"><span data-stu-id="5ca6f-108">Can contain</span></span>

|<span data-ttu-id="5ca6f-109">**Elemento**</span><span class="sxs-lookup"><span data-stu-id="5ca6f-109">**Element**</span></span>|<span data-ttu-id="5ca6f-110">**Content**</span><span class="sxs-lookup"><span data-stu-id="5ca6f-110">**Content**</span></span>|<span data-ttu-id="5ca6f-111">**Email**</span><span class="sxs-lookup"><span data-stu-id="5ca6f-111">**Mail**</span></span>|<span data-ttu-id="5ca6f-112">**TaskPane**</span><span class="sxs-lookup"><span data-stu-id="5ca6f-112">**TaskPane**</span></span>|
|:-----|:-----|:-----|:-----|
|[<span data-ttu-id="5ca6f-113">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="5ca6f-113">SourceLocation</span></span>](sourcelocation.md)|<span data-ttu-id="5ca6f-114">x</span><span class="sxs-lookup"><span data-stu-id="5ca6f-114">x</span></span>||<span data-ttu-id="5ca6f-115">x</span><span class="sxs-lookup"><span data-stu-id="5ca6f-115">x</span></span>|
|[<span data-ttu-id="5ca6f-116">RequestedWidth</span><span class="sxs-lookup"><span data-stu-id="5ca6f-116">RequestedWidth</span></span>](requestedwidth.md)|<span data-ttu-id="5ca6f-117">x</span><span class="sxs-lookup"><span data-stu-id="5ca6f-117">x</span></span>|||
|[<span data-ttu-id="5ca6f-118">RequestedHeight</span><span class="sxs-lookup"><span data-stu-id="5ca6f-118">RequestedHeight</span></span>](requestedheight.md)|<span data-ttu-id="5ca6f-119">x</span><span class="sxs-lookup"><span data-stu-id="5ca6f-119">x</span></span>|||

## <a name="remarks"></a><span data-ttu-id="5ca6f-120">Comentários</span><span class="sxs-lookup"><span data-stu-id="5ca6f-120">Remarks</span></span>

<span data-ttu-id="5ca6f-121">O local de origem e outras configurações no elemento **DefaultSettings** se aplicam apenas a suplementos de conteúdo e de painel de tarefas. Para suplementos de email, você especifica os locais padrão para os arquivos de origem e outras configurações padrão no elemento [FormSettings](formsettings.md).</span><span class="sxs-lookup"><span data-stu-id="5ca6f-121">The source location and other settings in the  **DefaultSettings** element apply only to content and task pane add-ins. For mail add-ins, you specify the default locations for source files and other default settings in the [FormSettings](formsettings.md) element.</span></span>

