---
title: Elemento RequestedHeight no arquivo de manifesto
description: ''
ms.date: 10/09/2018
ms.openlocfilehash: ea8c0403146f526b28eb20b8364fd210ac357baf
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 12/22/2018
ms.locfileid: "27433471"
---
# <a name="requestedheight-element"></a><span data-ttu-id="856f2-102">Elemento RequestedHeight</span><span class="sxs-lookup"><span data-stu-id="856f2-102">RequestedHeight element</span></span>

<span data-ttu-id="856f2-103">Especifica a altura inicial (em pixels) de um suplemento de conteúdo ou de email.</span><span class="sxs-lookup"><span data-stu-id="856f2-103">Specifies the initial height (in pixels) of a content add-in or mail add-in.</span></span> 

<span data-ttu-id="856f2-104">**Tipo de suplemento:** Conteúdo, Email</span><span class="sxs-lookup"><span data-stu-id="856f2-104">**Add-in type:** Content, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="856f2-105">Sintaxe</span><span class="sxs-lookup"><span data-stu-id="856f2-105">Syntax</span></span>

```XML
<RequestedHeight>integer</RequestedHeight>
```

## <a name="contained-in"></a><span data-ttu-id="856f2-106">Contido em</span><span class="sxs-lookup"><span data-stu-id="856f2-106">Contained in</span></span>

- <span data-ttu-id="856f2-107">[DefaultSettings](defaultsettings.md) (suplementos de conteúdo) com um valor entre 32 e 1000</span><span class="sxs-lookup"><span data-stu-id="856f2-107">[DefaultSettings](defaultsettings.md) (Content add-ins) with a value that can be between 32 and 1000</span></span>
- <span data-ttu-id="856f2-108">[DesktopSettings](desktopsettings.md) e [TabletSettings](tabletsettings.md) (suplementos de email ) com um valor entre 32 e 450</span><span class="sxs-lookup"><span data-stu-id="856f2-108">[DesktopSettings](desktopsettings.md) and [TabletSettings](tabletsettings.md) (Mail add-ins) with a value that can be between 32 and 450</span></span>
- <span data-ttu-id="856f2-109">[ExtensionPoint](extensionpoint.md) (suplementos contextuais de email) com um valor entre 140 e 450 para o ponto de extensão **DetectedEntity** e entre 32 e 450 para o ponto de extensão **CustomPane**</span><span class="sxs-lookup"><span data-stu-id="856f2-109">[ExtensionPoint](extensionpoint.md) (Contextual mail add-ins) with a value that can be between 140 and 450 for the **DetectedEntity** extension point and between 32 and 450 for the **CustomPane** extension point</span></span>