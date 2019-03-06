---
title: Elemento SupportsSharedFolders no arquivo de manifesto
description: ''
ms.date: 03/01/2019
localization_priority: Normal
ms.openlocfilehash: bfbce42c7d1aa5eefab40b528c5b622aa7d2d54f
ms.sourcegitcommit: 7ebd383f16ae5809bb6980a5f213b695d410e62c
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/06/2019
ms.locfileid: "30413612"
---
# <a name="supportssharedfolders-element"></a><span data-ttu-id="d3fad-102">Elemento SupportsSharedFolders</span><span class="sxs-lookup"><span data-stu-id="d3fad-102">SupportsSharedFolders element</span></span>

<span data-ttu-id="d3fad-103">Define se o suplemento do Outlook está disponível nos cenários de representante.</span><span class="sxs-lookup"><span data-stu-id="d3fad-103">Defines whether the Outlook add-in is available in delegate scenarios.</span></span> <span data-ttu-id="d3fad-104">O **SupportsSharedFolders** é um elemento filho de [DesktopFormFactor](desktopformfactor.md).</span><span class="sxs-lookup"><span data-stu-id="d3fad-104">The **SupportsSharedFolders** element is a child element of [DesktopFormFactor](desktopformfactor.md).</span></span> <span data-ttu-id="d3fad-105">Ele é definido como *false* por padrão.</span><span class="sxs-lookup"><span data-stu-id="d3fad-105">It is set to *false* by default.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="d3fad-106">O acesso de representante para suplementos do Outlook está atualmente em visualização e é suportado apenas em clientes que são executados no Exchange Online.</span><span class="sxs-lookup"><span data-stu-id="d3fad-106">Delegate access for Outlook add-ins is currently in preview and only supported in clients that run against Exchange Online.</span></span> <span data-ttu-id="d3fad-107">Os suplementos que usam esse elemento não podem ser publicados no AppSource nem implantados por meio da implantação centralizada.</span><span class="sxs-lookup"><span data-stu-id="d3fad-107">Add-ins that use this element cannot be published to AppSource or deployed via centralized deployment.</span></span>

<span data-ttu-id="d3fad-108">Veja a seguir um exemplo do elemento **SupportsSharedFolders**.</span><span class="sxs-lookup"><span data-stu-id="d3fad-108">The following is an example of the  **SupportsSharedFolders** element.</span></span>

```XML
<DesktopFormFactor>
  <FunctionFile resid="residDesktopFuncUrl" />
  <SupportsSharedFolders>true</SupportsSharedFolders>
  <ExtensionPoint xsi:type="MessageReadCommandSurface">
    <!-- configure selected extension point -->
  </ExtensionPoint>

  <!-- You can define more than one ExtensionPoint element as needed -->

</DesktopFormFactor>
```
