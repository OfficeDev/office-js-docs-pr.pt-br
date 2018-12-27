---
title: Elemento DefaultSettings no arquivo de manifesto
description: ''
ms.date: 10/09/2018
ms.openlocfilehash: 0c109d5d893cf9d3502f1cbf1724007f01e623e6
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 12/22/2018
ms.locfileid: "27433751"
---
# <a name="defaultsettings-element"></a><span data-ttu-id="633be-102">Elemento DefaultSettings</span><span class="sxs-lookup"><span data-stu-id="633be-102">DefaultSettings element</span></span>

<span data-ttu-id="633be-103">Especifica a localização de origem padrão e outras configurações padrão para o suplemento de conteúdo ou de painel de tarefas.</span><span class="sxs-lookup"><span data-stu-id="633be-103">Specifies the default source location and other default settings for your content or task pane add-in .</span></span>

<span data-ttu-id="633be-104">**Tipo de suplemento:** Conteúdo, Painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="633be-104">**Add-in type:** Content, Task pane</span></span>

## <a name="syntax"></a><span data-ttu-id="633be-105">Sintaxe</span><span class="sxs-lookup"><span data-stu-id="633be-105">Syntax</span></span>

```XML
<DefaultSettings>
  ...
</DefaultSettings>
```

## <a name="contained-in"></a><span data-ttu-id="633be-106">Contido em</span><span class="sxs-lookup"><span data-stu-id="633be-106">Contained in</span></span>

[<span data-ttu-id="633be-107">OfficeApp</span><span class="sxs-lookup"><span data-stu-id="633be-107">OfficeApp</span></span>](officeapp.md)

## <a name="can-contain"></a><span data-ttu-id="633be-108">Pode conter</span><span class="sxs-lookup"><span data-stu-id="633be-108">Can contain</span></span>

|<span data-ttu-id="633be-109">**Element**</span><span class="sxs-lookup"><span data-stu-id="633be-109">**Element**</span></span>|<span data-ttu-id="633be-110">**Conteúdo**</span><span class="sxs-lookup"><span data-stu-id="633be-110">**Content**</span></span>|<span data-ttu-id="633be-111">**Email**</span><span class="sxs-lookup"><span data-stu-id="633be-111">**Mail**</span></span>|<span data-ttu-id="633be-112">**TaskPane**</span><span class="sxs-lookup"><span data-stu-id="633be-112">**TaskPane**</span></span>|
|:-----|:-----|:-----|:-----|
|[<span data-ttu-id="633be-113">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="633be-113">SourceLocation</span></span>](sourcelocation.md)|<span data-ttu-id="633be-114">x</span><span class="sxs-lookup"><span data-stu-id="633be-114">x</span></span>||<span data-ttu-id="633be-115">x</span><span class="sxs-lookup"><span data-stu-id="633be-115">x</span></span>|
|[<span data-ttu-id="633be-116">RequestedWidth</span><span class="sxs-lookup"><span data-stu-id="633be-116">RequestedWidth</span></span>](requestedwidth.md)|<span data-ttu-id="633be-117">x</span><span class="sxs-lookup"><span data-stu-id="633be-117">x</span></span>|||
|[<span data-ttu-id="633be-118">RequestedHeight</span><span class="sxs-lookup"><span data-stu-id="633be-118">RequestedHeight</span></span>](requestedheight.md)|<span data-ttu-id="633be-119">x</span><span class="sxs-lookup"><span data-stu-id="633be-119">x</span></span>|||

## <a name="remarks"></a><span data-ttu-id="633be-120">Comentários</span><span class="sxs-lookup"><span data-stu-id="633be-120">Remarks</span></span>

<span data-ttu-id="633be-121">O local de origem e outras configurações no elemento **DefaultSettings** se aplicam apenas a suplementos de conteúdo e de painel de tarefas. Para suplementos de email, você especifica os locais padrão para os arquivos de origem e outras configurações padrão no elemento [FormSettings](formsettings.md).</span><span class="sxs-lookup"><span data-stu-id="633be-121">The source location and other settings in the  **DefaultSettings** element apply only to content and task pane add-ins. For mail add-ins, you specify the default locations for source files and other default settings in the [FormSettings](formsettings.md) element.</span></span>

