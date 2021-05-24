---
title: Elemento SourceLocation no arquivo de manifesto
description: O elemento SourceLocation especifica os locais do arquivo de origem para o Office Do-in.
ms.date: 05/12/2021
localization_priority: Normal
ms.openlocfilehash: 4dcd093db2f23220eaa34c0c81300c4994c1a697
ms.sourcegitcommit: 0d9fcdc2aeb160ff475fbe817425279267c7ff31
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/21/2021
ms.locfileid: "52590894"
---
# <a name="sourcelocation-element"></a><span data-ttu-id="78b66-103">Elemento SourceLocation</span><span class="sxs-lookup"><span data-stu-id="78b66-103">SourceLocation element</span></span>

<span data-ttu-id="78b66-104">Especifica os locais de arquivo de origem do seu Office como uma URL entre 1 e 2018 caracteres.</span><span class="sxs-lookup"><span data-stu-id="78b66-104">Specifies the source file locations for your Office Add-in as a URL between 1 and 2018 characters long.</span></span> <span data-ttu-id="78b66-105">O local de origem deve ser um endereço HTTPS, não um caminho de arquivo.</span><span class="sxs-lookup"><span data-stu-id="78b66-105">The source location must be an HTTPS address, not a file path.</span></span>

<span data-ttu-id="78b66-106">**Tipo de suplemento:** Conteúdo, Painel de tarefas, Email</span><span class="sxs-lookup"><span data-stu-id="78b66-106">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="78b66-107">Sintaxe</span><span class="sxs-lookup"><span data-stu-id="78b66-107">Syntax</span></span>

```XML
<SourceLocation DefaultValue="string" />
```

## <a name="contained-in"></a><span data-ttu-id="78b66-108">Contido em</span><span class="sxs-lookup"><span data-stu-id="78b66-108">Contained in</span></span>

- <span data-ttu-id="78b66-109">[DefaultSettings](defaultsettings.md) (suplementos de conteúdo e de painel de tarefas)</span><span class="sxs-lookup"><span data-stu-id="78b66-109">[DefaultSettings](defaultsettings.md) (Content and task pane add-ins)</span></span>
- <span data-ttu-id="78b66-110">[FormSettings](formsettings.md) (suplementos de email)</span><span class="sxs-lookup"><span data-stu-id="78b66-110">[FormSettings](formsettings.md) (Mail add-ins)</span></span>
- <span data-ttu-id="78b66-111">[ExtensionPoint](extensionpoint.md) (Contextual e LaunchEvent mail add-ins)</span><span class="sxs-lookup"><span data-stu-id="78b66-111">[ExtensionPoint](extensionpoint.md) (Contextual and LaunchEvent mail add-ins)</span></span>

## <a name="can-contain"></a><span data-ttu-id="78b66-112">Pode conter</span><span class="sxs-lookup"><span data-stu-id="78b66-112">Can contain</span></span>

[<span data-ttu-id="78b66-113">Override</span><span class="sxs-lookup"><span data-stu-id="78b66-113">Override</span></span>](override.md)

## <a name="attributes"></a><span data-ttu-id="78b66-114">Atributos</span><span class="sxs-lookup"><span data-stu-id="78b66-114">Attributes</span></span>

|<span data-ttu-id="78b66-115">Atributo</span><span class="sxs-lookup"><span data-stu-id="78b66-115">Attribute</span></span>|<span data-ttu-id="78b66-116">Tipo</span><span class="sxs-lookup"><span data-stu-id="78b66-116">Type</span></span>|<span data-ttu-id="78b66-117">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="78b66-117">Required</span></span>|<span data-ttu-id="78b66-118">Descrição</span><span class="sxs-lookup"><span data-stu-id="78b66-118">Description</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="78b66-119">DefaultValue</span><span class="sxs-lookup"><span data-stu-id="78b66-119">DefaultValue</span></span>|<span data-ttu-id="78b66-120">URL</span><span class="sxs-lookup"><span data-stu-id="78b66-120">URL</span></span>|<span data-ttu-id="78b66-121">obrigatório</span><span class="sxs-lookup"><span data-stu-id="78b66-121">required</span></span>|<span data-ttu-id="78b66-122">Especifica o valor padrão para essa configuração para a localidade especificada no elemento [DefaultLocale](defaultlocale.md).</span><span class="sxs-lookup"><span data-stu-id="78b66-122">Specifies the default value for this setting for the locale specified in the [DefaultLocale](defaultlocale.md) element.</span></span>|
