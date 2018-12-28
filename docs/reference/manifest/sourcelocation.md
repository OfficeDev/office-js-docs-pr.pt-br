---
title: Elemento SourceLocation no arquivo de manifesto
description: ''
ms.date: 10/09/2018
ms.openlocfilehash: dc432ebb9482e8e9b8be5d90a838357ccf519ad3
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 12/22/2018
ms.locfileid: "27433513"
---
# <a name="sourcelocation-element"></a><span data-ttu-id="81415-102">Elemento SourceLocation</span><span class="sxs-lookup"><span data-stu-id="81415-102">SourceLocation element</span></span>

<span data-ttu-id="81415-p101">Especifica o local de origem do arquivo para o Suplemento do Office como uma URL que contém entre 1 e 2.018 caracteres de comprimento. O local de origem deve ser um endereço HTTPS, não um caminho de arquivo.</span><span class="sxs-lookup"><span data-stu-id="81415-p101">Specifies the source file location(s) for your Office Add-in as a URL between 1 and 2018 characters long. The source location must be an HTTPS address, not a file path.</span></span>

<span data-ttu-id="81415-105">**Tipo de suplemento:** Conteúdo, Painel de tarefas, Email</span><span class="sxs-lookup"><span data-stu-id="81415-105">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="81415-106">Sintaxe</span><span class="sxs-lookup"><span data-stu-id="81415-106">Syntax</span></span>

```XML
<SourceLocation DefaultValue="string" />
```

## <a name="contained-in"></a><span data-ttu-id="81415-107">Contido em</span><span class="sxs-lookup"><span data-stu-id="81415-107">Contained in</span></span>

- <span data-ttu-id="81415-108">[DefaultSettings](defaultsettings.md) (suplementos de conteúdo e de painel de tarefas)</span><span class="sxs-lookup"><span data-stu-id="81415-108">[DefaultSettings](defaultsettings.md) (Content and task pane add-ins)</span></span>
- <span data-ttu-id="81415-109">[FormSettings](formsettings.md) (suplementos de email)</span><span class="sxs-lookup"><span data-stu-id="81415-109">[FormSettings](formsettings.md) (Mail add-ins)</span></span>
- <span data-ttu-id="81415-110">[ExtensionPoint](extensionpoint.md) (suplementos contextuais de email)</span><span class="sxs-lookup"><span data-stu-id="81415-110">[ExtensionPoint](extensionpoint.md) (Contextual mail add-ins)</span></span>

## <a name="can-contain"></a><span data-ttu-id="81415-111">Pode conter</span><span class="sxs-lookup"><span data-stu-id="81415-111">Can contain</span></span>

[<span data-ttu-id="81415-112">Override</span><span class="sxs-lookup"><span data-stu-id="81415-112">Override</span></span>](override.md)

## <a name="attributes"></a><span data-ttu-id="81415-113">Atributos</span><span class="sxs-lookup"><span data-stu-id="81415-113">Attributes</span></span>

|<span data-ttu-id="81415-114">**Atributo**</span><span class="sxs-lookup"><span data-stu-id="81415-114">**Attribute**</span></span>|<span data-ttu-id="81415-115">**Tipo**</span><span class="sxs-lookup"><span data-stu-id="81415-115">**Type**</span></span>|<span data-ttu-id="81415-116">**Obrigatório**</span><span class="sxs-lookup"><span data-stu-id="81415-116">**Required**</span></span>|<span data-ttu-id="81415-117">**Descrição**</span><span class="sxs-lookup"><span data-stu-id="81415-117">**Description**</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="81415-118">DefaultValue</span><span class="sxs-lookup"><span data-stu-id="81415-118">DefaultValue</span></span>|<span data-ttu-id="81415-119">URL</span><span class="sxs-lookup"><span data-stu-id="81415-119">URL</span></span>|<span data-ttu-id="81415-120">obrigatório</span><span class="sxs-lookup"><span data-stu-id="81415-120">required</span></span>|<span data-ttu-id="81415-121">Especifica o valor padrão para essa configuração para a localidade especificada no elemento [DefaultLocale](defaultlocale.md).</span><span class="sxs-lookup"><span data-stu-id="81415-121">Specifies the default value for this setting for the locale specified in the [DefaultLocale](defaultlocale.md) element.</span></span>|
