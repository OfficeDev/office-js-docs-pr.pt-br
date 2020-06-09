---
title: Elemento RequestedHeight no arquivo de manifesto
description: O elemento RequestedHeight especifica a altura inicial (em pixels) de um suplemento de conteúdo ou email.
ms.date: 05/14/2020
localization_priority: Normal
ms.openlocfilehash: 44675918a4208683f442fe8a6e8f4f906f484571
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 06/08/2020
ms.locfileid: "44611726"
---
# <a name="requestedheight-element"></a><span data-ttu-id="b18a7-103">Elemento RequestedHeight</span><span class="sxs-lookup"><span data-stu-id="b18a7-103">RequestedHeight element</span></span>

<span data-ttu-id="b18a7-104">Especifica a altura inicial (em pixels) de um suplemento de conteúdo ou de email.</span><span class="sxs-lookup"><span data-stu-id="b18a7-104">Specifies the initial height (in pixels) of a content add-in or mail add-in.</span></span>

<span data-ttu-id="b18a7-105">**Tipo de suplemento:** Conteúdo, Email</span><span class="sxs-lookup"><span data-stu-id="b18a7-105">**Add-in type:** Content, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="b18a7-106">Sintaxe</span><span class="sxs-lookup"><span data-stu-id="b18a7-106">Syntax</span></span>

```XML
<RequestedHeight>integer</RequestedHeight>
```

## <a name="contained-in"></a><span data-ttu-id="b18a7-107">Contido em</span><span class="sxs-lookup"><span data-stu-id="b18a7-107">Contained in</span></span>

- <span data-ttu-id="b18a7-108">[DefaultSettings](defaultsettings.md) (suplementos de conteúdo) com um valor entre 32 e 1000</span><span class="sxs-lookup"><span data-stu-id="b18a7-108">[DefaultSettings](defaultsettings.md) (Content add-ins) with a value that can be between 32 and 1000</span></span>
- <span data-ttu-id="b18a7-109">[DesktopSettings](desktopsettings.md) e [TabletSettings](tabletsettings.md) (suplementos de email ) com um valor entre 32 e 450</span><span class="sxs-lookup"><span data-stu-id="b18a7-109">[DesktopSettings](desktopsettings.md) and [TabletSettings](tabletsettings.md) (Mail add-ins) with a value that can be between 32 and 450</span></span>
- <span data-ttu-id="b18a7-110">[ExtensionPoint](extensionpoint.md) (suplementos de email contextuais) com um valor que pode ser entre 140 e 450 para o ponto de extensão **DetectedEntity** e entre 32 e 450 para o [ponto de extensão **CustomPane** (preterido)](https://developer.microsoft.com/outlook/blogs/make-your-add-ins-available-in-the-office-ribbon/)</span><span class="sxs-lookup"><span data-stu-id="b18a7-110">[ExtensionPoint](extensionpoint.md) (Contextual mail add-ins) with a value that can be between 140 and 450 for the **DetectedEntity** extension point and between 32 and 450 for the [**CustomPane** extension point (deprecated)](https://developer.microsoft.com/outlook/blogs/make-your-add-ins-available-in-the-office-ribbon/)</span></span>
