---
title: Elemento DefaultSettings no arquivo de manifesto
description: Especifica a localização de origem padrão e outras configurações padrão para o suplemento de conteúdo ou de painel de tarefas.
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: a9711fb44390bcbda8979b8018eed1318c5579bc
ms.sourcegitcommit: cc6886b47c84ac37a3c957ff85dd0ed526ca5e43
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 08/12/2020
ms.locfileid: "46641463"
---
# <a name="defaultsettings-element"></a><span data-ttu-id="8443e-103">Elemento DefaultSettings</span><span class="sxs-lookup"><span data-stu-id="8443e-103">DefaultSettings element</span></span>

<span data-ttu-id="8443e-104">Especifica a localização de origem padrão e outras configurações padrão para o suplemento de conteúdo ou de painel de tarefas.</span><span class="sxs-lookup"><span data-stu-id="8443e-104">Specifies the default source location and other default settings for your content or task pane add-in.</span></span>

<span data-ttu-id="8443e-105">**Tipo de suplemento:** Conteúdo, Painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="8443e-105">**Add-in type:** Content, Task pane</span></span>

## <a name="syntax"></a><span data-ttu-id="8443e-106">Sintaxe</span><span class="sxs-lookup"><span data-stu-id="8443e-106">Syntax</span></span>

```XML
<DefaultSettings>
  ...
</DefaultSettings>
```

## <a name="contained-in"></a><span data-ttu-id="8443e-107">Contido em</span><span class="sxs-lookup"><span data-stu-id="8443e-107">Contained in</span></span>

[<span data-ttu-id="8443e-108">OfficeApp</span><span class="sxs-lookup"><span data-stu-id="8443e-108">OfficeApp</span></span>](officeapp.md)

## <a name="can-contain"></a><span data-ttu-id="8443e-109">Pode conter</span><span class="sxs-lookup"><span data-stu-id="8443e-109">Can contain</span></span>

|<span data-ttu-id="8443e-110">Elemento</span><span class="sxs-lookup"><span data-stu-id="8443e-110">Element</span></span>|<span data-ttu-id="8443e-111">Conteúdo</span><span class="sxs-lookup"><span data-stu-id="8443e-111">Content</span></span>|<span data-ttu-id="8443e-112">Email</span><span class="sxs-lookup"><span data-stu-id="8443e-112">Mail</span></span>|<span data-ttu-id="8443e-113">TaskPane</span><span class="sxs-lookup"><span data-stu-id="8443e-113">TaskPane</span></span>|
|:-----|:-----|:-----|:-----|
|[<span data-ttu-id="8443e-114">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="8443e-114">SourceLocation</span></span>](sourcelocation.md)|<span data-ttu-id="8443e-115">x</span><span class="sxs-lookup"><span data-stu-id="8443e-115">x</span></span>||<span data-ttu-id="8443e-116">x</span><span class="sxs-lookup"><span data-stu-id="8443e-116">x</span></span>|
|[<span data-ttu-id="8443e-117">RequestedWidth</span><span class="sxs-lookup"><span data-stu-id="8443e-117">RequestedWidth</span></span>](requestedwidth.md)|<span data-ttu-id="8443e-118">x</span><span class="sxs-lookup"><span data-stu-id="8443e-118">x</span></span>|||
|[<span data-ttu-id="8443e-119">RequestedHeight</span><span class="sxs-lookup"><span data-stu-id="8443e-119">RequestedHeight</span></span>](requestedheight.md)|<span data-ttu-id="8443e-120">x</span><span class="sxs-lookup"><span data-stu-id="8443e-120">x</span></span>|||

## <a name="remarks"></a><span data-ttu-id="8443e-121">Comentários</span><span class="sxs-lookup"><span data-stu-id="8443e-121">Remarks</span></span>

<span data-ttu-id="8443e-122">O local de origem e outras configurações no elemento **DefaultSettings** só se aplicam a suplementos de conteúdo e de painel de tarefas. Para suplementos de email, você especifica os locais padrão para arquivos de origem e outras configurações padrão no elemento [FormSettings](formsettings.md) .</span><span class="sxs-lookup"><span data-stu-id="8443e-122">The source location and other settings in the **DefaultSettings** element apply only to content and task pane add-ins. For mail add-ins, you specify the default locations for source files and other default settings in the [FormSettings](formsettings.md) element.</span></span>
