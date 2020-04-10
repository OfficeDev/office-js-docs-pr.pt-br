---
title: Elemento RequestedHeight no arquivo de manifesto
description: O elemento RequestedHeight especifica a altura inicial (em pixels) de um suplemento de conteúdo ou email.
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: 5f4c3ca1ff39cc3150249fbc824b0db76f6b8a85
ms.sourcegitcommit: c6e3bfd3deb77982d0b7082afd6a48678e96e1c3
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 04/10/2020
ms.locfileid: "43215037"
---
# <a name="requestedheight-element"></a><span data-ttu-id="8cb37-103">Elemento RequestedHeight</span><span class="sxs-lookup"><span data-stu-id="8cb37-103">RequestedHeight element</span></span>

<span data-ttu-id="8cb37-104">Especifica a altura inicial (em pixels) de um suplemento de conteúdo ou de email.</span><span class="sxs-lookup"><span data-stu-id="8cb37-104">Specifies the initial height (in pixels) of a content add-in or mail add-in.</span></span>

<span data-ttu-id="8cb37-105">**Tipo de suplemento:** Conteúdo, Email</span><span class="sxs-lookup"><span data-stu-id="8cb37-105">**Add-in type:** Content, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="8cb37-106">Sintaxe</span><span class="sxs-lookup"><span data-stu-id="8cb37-106">Syntax</span></span>

```XML
<RequestedHeight>integer</RequestedHeight>
```

## <a name="contained-in"></a><span data-ttu-id="8cb37-107">Contido em</span><span class="sxs-lookup"><span data-stu-id="8cb37-107">Contained in</span></span>

- <span data-ttu-id="8cb37-108">[DefaultSettings](defaultsettings.md) (suplementos de conteúdo) com um valor entre 32 e 1000</span><span class="sxs-lookup"><span data-stu-id="8cb37-108">[DefaultSettings](defaultsettings.md) (Content add-ins) with a value that can be between 32 and 1000</span></span>
- <span data-ttu-id="8cb37-109">[DesktopSettings](desktopsettings.md) e [TabletSettings](tabletsettings.md) (suplementos de email ) com um valor entre 32 e 450</span><span class="sxs-lookup"><span data-stu-id="8cb37-109">[DesktopSettings](desktopsettings.md) and [TabletSettings](tabletsettings.md) (Mail add-ins) with a value that can be between 32 and 450</span></span>
- <span data-ttu-id="8cb37-110">[ExtensionPoint](extensionpoint.md) (suplementos contextuais de email) com um valor entre 140 e 450 para o ponto de extensão **DetectedEntity** e entre 32 e 450 para o ponto de extensão **CustomPane**</span><span class="sxs-lookup"><span data-stu-id="8cb37-110">[ExtensionPoint](extensionpoint.md) (Contextual mail add-ins) with a value that can be between 140 and 450 for the **DetectedEntity** extension point and between 32 and 450 for the **CustomPane** extension point</span></span>
