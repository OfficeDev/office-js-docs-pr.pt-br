---
title: Elemento SourceLocation no arquivo de manifesto
description: O elemento SourceLocation especifica os locais do arquivo de origem para o suplemento do Office.
ms.date: 05/12/2020
localization_priority: Normal
ms.openlocfilehash: 9af2337263314bec5ce04eb0d22626ab368c19ef
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 06/08/2020
ms.locfileid: "44608723"
---
# <a name="sourcelocation-element"></a><span data-ttu-id="a47d9-103">Elemento SourceLocation</span><span class="sxs-lookup"><span data-stu-id="a47d9-103">SourceLocation element</span></span>

<span data-ttu-id="a47d9-104">Especifica os locais do arquivo de origem para o suplemento do Office como uma URL entre 1 e 2018 caracteres de comprimento.</span><span class="sxs-lookup"><span data-stu-id="a47d9-104">Specifies the source file locations for your Office Add-in as a URL between 1 and 2018 characters long.</span></span> <span data-ttu-id="a47d9-105">O local de origem deve ser um endereço HTTPS, não um caminho de arquivo.</span><span class="sxs-lookup"><span data-stu-id="a47d9-105">The source location must be an HTTPS address, not a file path.</span></span>

<span data-ttu-id="a47d9-106">**Tipo de suplemento:** Conteúdo, Painel de tarefas, Email</span><span class="sxs-lookup"><span data-stu-id="a47d9-106">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="a47d9-107">Sintaxe</span><span class="sxs-lookup"><span data-stu-id="a47d9-107">Syntax</span></span>

```XML
<SourceLocation DefaultValue="string" />
```

## <a name="contained-in"></a><span data-ttu-id="a47d9-108">Contido em</span><span class="sxs-lookup"><span data-stu-id="a47d9-108">Contained in</span></span>

- <span data-ttu-id="a47d9-109">[DefaultSettings](defaultsettings.md) (suplementos de conteúdo e de painel de tarefas)</span><span class="sxs-lookup"><span data-stu-id="a47d9-109">[DefaultSettings](defaultsettings.md) (Content and task pane add-ins)</span></span>
- <span data-ttu-id="a47d9-110">[FormSettings](formsettings.md) (suplementos de email)</span><span class="sxs-lookup"><span data-stu-id="a47d9-110">[FormSettings](formsettings.md) (Mail add-ins)</span></span>
- <span data-ttu-id="a47d9-111">[ExtensionPoint](extensionpoint.md) (contextuais e LaunchEvent (Visualizar) suplementos de email)</span><span class="sxs-lookup"><span data-stu-id="a47d9-111">[ExtensionPoint](extensionpoint.md) (Contextual and LaunchEvent (preview) mail add-ins)</span></span>

## <a name="can-contain"></a><span data-ttu-id="a47d9-112">Pode conter</span><span class="sxs-lookup"><span data-stu-id="a47d9-112">Can contain</span></span>

[<span data-ttu-id="a47d9-113">Override</span><span class="sxs-lookup"><span data-stu-id="a47d9-113">Override</span></span>](override.md)

## <a name="attributes"></a><span data-ttu-id="a47d9-114">Atributos</span><span class="sxs-lookup"><span data-stu-id="a47d9-114">Attributes</span></span>

|<span data-ttu-id="a47d9-115">**Atributo**</span><span class="sxs-lookup"><span data-stu-id="a47d9-115">**Attribute**</span></span>|<span data-ttu-id="a47d9-116">**Tipo**</span><span class="sxs-lookup"><span data-stu-id="a47d9-116">**Type**</span></span>|<span data-ttu-id="a47d9-117">**Obrigatório**</span><span class="sxs-lookup"><span data-stu-id="a47d9-117">**Required**</span></span>|<span data-ttu-id="a47d9-118">**Descrição**</span><span class="sxs-lookup"><span data-stu-id="a47d9-118">**Description**</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="a47d9-119">DefaultValue</span><span class="sxs-lookup"><span data-stu-id="a47d9-119">DefaultValue</span></span>|<span data-ttu-id="a47d9-120">URL</span><span class="sxs-lookup"><span data-stu-id="a47d9-120">URL</span></span>|<span data-ttu-id="a47d9-121">obrigatório</span><span class="sxs-lookup"><span data-stu-id="a47d9-121">required</span></span>|<span data-ttu-id="a47d9-122">Especifica o valor padrão para essa configuração para a localidade especificada no elemento [DefaultLocale](defaultlocale.md).</span><span class="sxs-lookup"><span data-stu-id="a47d9-122">Specifies the default value for this setting for the locale specified in the [DefaultLocale](defaultlocale.md) element.</span></span>|
