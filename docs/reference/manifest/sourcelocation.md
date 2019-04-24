---
title: Elemento SourceLocation no arquivo de manifesto
description: ''
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: 7544e2bae480b9431c8912533ea1b761132a355e
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 04/24/2019
ms.locfileid: "32451973"
---
# <a name="sourcelocation-element"></a><span data-ttu-id="67a0e-102">Elemento SourceLocation</span><span class="sxs-lookup"><span data-stu-id="67a0e-102">SourceLocation element</span></span>

<span data-ttu-id="67a0e-p101">Especifica o local de origem do arquivo para o Suplemento do Office como uma URL que contém entre 1 e 2.018 caracteres de comprimento. O local de origem deve ser um endereço HTTPS, não um caminho de arquivo.</span><span class="sxs-lookup"><span data-stu-id="67a0e-p101">Specifies the source file location(s) for your Office Add-in as a URL between 1 and 2018 characters long. The source location must be an HTTPS address, not a file path.</span></span>

<span data-ttu-id="67a0e-105">**Tipo de suplemento:** Conteúdo, Painel de tarefas, Email</span><span class="sxs-lookup"><span data-stu-id="67a0e-105">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="67a0e-106">Sintaxe</span><span class="sxs-lookup"><span data-stu-id="67a0e-106">Syntax</span></span>

```XML
<SourceLocation DefaultValue="string" />
```

## <a name="contained-in"></a><span data-ttu-id="67a0e-107">Contido em</span><span class="sxs-lookup"><span data-stu-id="67a0e-107">Contained in</span></span>

- <span data-ttu-id="67a0e-108">[DefaultSettings](defaultsettings.md) (suplementos de conteúdo e de painel de tarefas)</span><span class="sxs-lookup"><span data-stu-id="67a0e-108">[DefaultSettings](defaultsettings.md) (Content and task pane add-ins)</span></span>
- <span data-ttu-id="67a0e-109">[FormSettings](formsettings.md) (suplementos de email)</span><span class="sxs-lookup"><span data-stu-id="67a0e-109">[FormSettings](formsettings.md) (Mail add-ins)</span></span>
- <span data-ttu-id="67a0e-110">[ExtensionPoint](extensionpoint.md) (suplementos contextuais de email)</span><span class="sxs-lookup"><span data-stu-id="67a0e-110">[ExtensionPoint](extensionpoint.md) (Contextual mail add-ins)</span></span>

## <a name="can-contain"></a><span data-ttu-id="67a0e-111">Pode conter</span><span class="sxs-lookup"><span data-stu-id="67a0e-111">Can contain</span></span>

[<span data-ttu-id="67a0e-112">Override</span><span class="sxs-lookup"><span data-stu-id="67a0e-112">Override</span></span>](override.md)

## <a name="attributes"></a><span data-ttu-id="67a0e-113">Atributos</span><span class="sxs-lookup"><span data-stu-id="67a0e-113">Attributes</span></span>

|<span data-ttu-id="67a0e-114">**Atributo**</span><span class="sxs-lookup"><span data-stu-id="67a0e-114">**Attribute**</span></span>|<span data-ttu-id="67a0e-115">**Tipo**</span><span class="sxs-lookup"><span data-stu-id="67a0e-115">**Type**</span></span>|<span data-ttu-id="67a0e-116">**Obrigatório**</span><span class="sxs-lookup"><span data-stu-id="67a0e-116">**Required**</span></span>|<span data-ttu-id="67a0e-117">**Descrição**</span><span class="sxs-lookup"><span data-stu-id="67a0e-117">**Description**</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="67a0e-118">DefaultValue</span><span class="sxs-lookup"><span data-stu-id="67a0e-118">DefaultValue</span></span>|<span data-ttu-id="67a0e-119">URL</span><span class="sxs-lookup"><span data-stu-id="67a0e-119">URL</span></span>|<span data-ttu-id="67a0e-120">obrigatório</span><span class="sxs-lookup"><span data-stu-id="67a0e-120">required</span></span>|<span data-ttu-id="67a0e-121">Especifica o valor padrão para essa configuração para a localidade especificada no elemento [DefaultLocale](defaultlocale.md).</span><span class="sxs-lookup"><span data-stu-id="67a0e-121">Specifies the default value for this setting for the locale specified in the [DefaultLocale](defaultlocale.md) element.</span></span>|
