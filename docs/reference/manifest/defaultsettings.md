---
title: Elemento DefaultSettings no arquivo de manifesto
description: ''
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: 824c575b39a99c6028ffd603390d2b41ee0ad7dd
ms.sourcegitcommit: 5d29801180f6939ec10efb778d2311be67d8b9f1
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 02/27/2020
ms.locfileid: "42324881"
---
# <a name="defaultsettings-element"></a><span data-ttu-id="8e9a3-102">Elemento DefaultSettings</span><span class="sxs-lookup"><span data-stu-id="8e9a3-102">DefaultSettings element</span></span>

<span data-ttu-id="8e9a3-103">Especifica a localização de origem padrão e outras configurações padrão para o suplemento de conteúdo ou de painel de tarefas.</span><span class="sxs-lookup"><span data-stu-id="8e9a3-103">Specifies the default source location and other default settings for your content or task pane add-in.</span></span>

<span data-ttu-id="8e9a3-104">**Tipo de suplemento:** Conteúdo, Painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="8e9a3-104">**Add-in type:** Content, Task pane</span></span>

## <a name="syntax"></a><span data-ttu-id="8e9a3-105">Sintaxe</span><span class="sxs-lookup"><span data-stu-id="8e9a3-105">Syntax</span></span>

```XML
<DefaultSettings>
  ...
</DefaultSettings>
```

## <a name="contained-in"></a><span data-ttu-id="8e9a3-106">Contido em</span><span class="sxs-lookup"><span data-stu-id="8e9a3-106">Contained in</span></span>

[<span data-ttu-id="8e9a3-107">OfficeApp</span><span class="sxs-lookup"><span data-stu-id="8e9a3-107">OfficeApp</span></span>](officeapp.md)

## <a name="can-contain"></a><span data-ttu-id="8e9a3-108">Pode conter</span><span class="sxs-lookup"><span data-stu-id="8e9a3-108">Can contain</span></span>

|<span data-ttu-id="8e9a3-109">**Elemento**</span><span class="sxs-lookup"><span data-stu-id="8e9a3-109">**Element**</span></span>|<span data-ttu-id="8e9a3-110">**Content**</span><span class="sxs-lookup"><span data-stu-id="8e9a3-110">**Content**</span></span>|<span data-ttu-id="8e9a3-111">**Email**</span><span class="sxs-lookup"><span data-stu-id="8e9a3-111">**Mail**</span></span>|<span data-ttu-id="8e9a3-112">**TaskPane**</span><span class="sxs-lookup"><span data-stu-id="8e9a3-112">**TaskPane**</span></span>|
|:-----|:-----|:-----|:-----|
|[<span data-ttu-id="8e9a3-113">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="8e9a3-113">SourceLocation</span></span>](sourcelocation.md)|<span data-ttu-id="8e9a3-114">x</span><span class="sxs-lookup"><span data-stu-id="8e9a3-114">x</span></span>||<span data-ttu-id="8e9a3-115">x</span><span class="sxs-lookup"><span data-stu-id="8e9a3-115">x</span></span>|
|[<span data-ttu-id="8e9a3-116">RequestedWidth</span><span class="sxs-lookup"><span data-stu-id="8e9a3-116">RequestedWidth</span></span>](requestedwidth.md)|<span data-ttu-id="8e9a3-117">x</span><span class="sxs-lookup"><span data-stu-id="8e9a3-117">x</span></span>|||
|[<span data-ttu-id="8e9a3-118">RequestedHeight</span><span class="sxs-lookup"><span data-stu-id="8e9a3-118">RequestedHeight</span></span>](requestedheight.md)|<span data-ttu-id="8e9a3-119">x</span><span class="sxs-lookup"><span data-stu-id="8e9a3-119">x</span></span>|||

## <a name="remarks"></a><span data-ttu-id="8e9a3-120">Comentários</span><span class="sxs-lookup"><span data-stu-id="8e9a3-120">Remarks</span></span>

<span data-ttu-id="8e9a3-121">O local de origem e outras configurações no elemento **DefaultSettings** só se aplicam a suplementos de conteúdo e de painel de tarefas. Para suplementos de email, você especifica os locais padrão para arquivos de origem e outras configurações padrão no elemento [FormSettings](formsettings.md) .</span><span class="sxs-lookup"><span data-stu-id="8e9a3-121">The source location and other settings in the **DefaultSettings** element apply only to content and task pane add-ins. For mail add-ins, you specify the default locations for source files and other default settings in the [FormSettings](formsettings.md) element.</span></span>

