---
title: Elemento RequestedHeight no arquivo de manifesto
description: O elemento RequestedHeight especifica a altura inicial (em pixels) de um suplemento de conteúdo ou email.
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: 853d12baf290167f3e6a635201e8b5d1d0e35a51
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/17/2020
ms.locfileid: "42720453"
---
# <a name="requestedheight-element"></a><span data-ttu-id="b87b7-103">Elemento RequestedHeight</span><span class="sxs-lookup"><span data-stu-id="b87b7-103">RequestedHeight element</span></span>

<span data-ttu-id="b87b7-104">Especifica a altura inicial (em pixels) de um suplemento de conteúdo ou de email.</span><span class="sxs-lookup"><span data-stu-id="b87b7-104">Specifies the initial height (in pixels) of a content add-in or mail add-in.</span></span> 

<span data-ttu-id="b87b7-105">**Tipo de suplemento:** Conteúdo, Email</span><span class="sxs-lookup"><span data-stu-id="b87b7-105">**Add-in type:** Content, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="b87b7-106">Sintaxe</span><span class="sxs-lookup"><span data-stu-id="b87b7-106">Syntax</span></span>

```XML
<RequestedHeight>integer</RequestedHeight>
```

## <a name="contained-in"></a><span data-ttu-id="b87b7-107">Contido em</span><span class="sxs-lookup"><span data-stu-id="b87b7-107">Contained in</span></span>

- <span data-ttu-id="b87b7-108">[DefaultSettings](defaultsettings.md) (suplementos de conteúdo) com um valor entre 32 e 1000</span><span class="sxs-lookup"><span data-stu-id="b87b7-108">[DefaultSettings](defaultsettings.md) (Content add-ins) with a value that can be between 32 and 1000</span></span>
- <span data-ttu-id="b87b7-109">[DesktopSettings](desktopsettings.md) e [TabletSettings](tabletsettings.md) (suplementos de email ) com um valor entre 32 e 450</span><span class="sxs-lookup"><span data-stu-id="b87b7-109">[DesktopSettings](desktopsettings.md) and [TabletSettings](tabletsettings.md) (Mail add-ins) with a value that can be between 32 and 450</span></span>
- <span data-ttu-id="b87b7-110">[ExtensionPoint](extensionpoint.md) (suplementos contextuais de email) com um valor entre 140 e 450 para o ponto de extensão **DetectedEntity** e entre 32 e 450 para o ponto de extensão **CustomPane**</span><span class="sxs-lookup"><span data-stu-id="b87b7-110">[ExtensionPoint](extensionpoint.md) (Contextual mail add-ins) with a value that can be between 140 and 450 for the **DetectedEntity** extension point and between 32 and 450 for the **CustomPane** extension point</span></span>
