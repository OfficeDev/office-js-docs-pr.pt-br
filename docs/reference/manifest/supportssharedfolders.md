---
title: Elemento SupportsSharedFolders no arquivo de manifesto
description: ''
ms.date: 10/09/2018
ms.openlocfilehash: 776d44ec66c4e27a72e5487051bed1edf4b3dcaf
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 12/22/2018
ms.locfileid: "27432680"
---
# <a name="supportssharedfolders-element"></a><span data-ttu-id="b3c7e-102">Elemento SupportsSharedFolders</span><span class="sxs-lookup"><span data-stu-id="b3c7e-102">SupportsSharedFolders element</span></span>

<span data-ttu-id="b3c7e-103">Define se o suplemento do Outlook está disponível nos cenários de representante.</span><span class="sxs-lookup"><span data-stu-id="b3c7e-103">It defines whether the add-in is available in delegate scenarios.</span></span> <span data-ttu-id="b3c7e-104">O **SupportsSharedFolders** é um elemento filho de [DesktopFormFactor](desktopformfactor.md).</span><span class="sxs-lookup"><span data-stu-id="b3c7e-104">The **ExtensionPoint** element is a child element of [AllFormFactors, DesktopFormFactor or MobileFormFactor](desktopformfactor.md).</span></span> <span data-ttu-id="b3c7e-105">Ele é definido como *false* por padrão.</span><span class="sxs-lookup"><span data-stu-id="b3c7e-105">It is set to *false* by default.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="b3c7e-106">Esse elemento só está disponível no [conjunto de requisitos de versão prévia de suplementos do Outlook](../objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) em comparação com o Exchange Online.</span><span class="sxs-lookup"><span data-stu-id="b3c7e-106">This element is only available in the [Outlook add-ins Preview requirement set](../objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) against Exchange Online.</span></span> <span data-ttu-id="b3c7e-107">Os suplementos que usam esse elemento não podem ser publicados no AppSource nem implantados por meio da implantação centralizada.</span><span class="sxs-lookup"><span data-stu-id="b3c7e-107">Add-ins that use this element cannot be published to AppSource or deployed via centralized deployment.</span></span>

<span data-ttu-id="b3c7e-108">Veja a seguir um exemplo do elemento **SupportsSharedFolders**.</span><span class="sxs-lookup"><span data-stu-id="b3c7e-108">The following is an example of the **FunctionFile** element.</span></span>

```XML
<DesktopFormFactor>
  <FunctionFile resid="residDesktopFuncUrl" />
  <SupportsSharedFolders>true</SupportsSharedFolders>
  <ExtensionPoint xsi:type="PrimaryCommandSurface">
    <!-- information about this extension point -->
  </ExtensionPoint>

  <!-- You can define more than one ExtensionPoint element as needed -->

</DesktopFormFactor>
```
