---
title: Elemento SupportsSharedFolders no arquivo de manifesto
description: ''
ms.date: 04/02/2019
localization_priority: Normal
ms.openlocfilehash: 976f8ba00f6ac9ac32def56933af1077527b7e9c
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 04/24/2019
ms.locfileid: "32452036"
---
# <a name="supportssharedfolders-element"></a><span data-ttu-id="0ecec-102">Elemento SupportsSharedFolders</span><span class="sxs-lookup"><span data-stu-id="0ecec-102">SupportsSharedFolders element</span></span>

<span data-ttu-id="0ecec-103">Define se o suplemento do Outlook está disponível nos cenários de representante.</span><span class="sxs-lookup"><span data-stu-id="0ecec-103">Defines whether the Outlook add-in is available in delegate scenarios.</span></span> <span data-ttu-id="0ecec-104">O **SupportsSharedFolders** é um elemento filho de [DesktopFormFactor](desktopformfactor.md).</span><span class="sxs-lookup"><span data-stu-id="0ecec-104">The **SupportsSharedFolders** element is a child element of [DesktopFormFactor](desktopformfactor.md).</span></span> <span data-ttu-id="0ecec-105">Ele é definido como *false* por padrão.</span><span class="sxs-lookup"><span data-stu-id="0ecec-105">It is set to *false* by default.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="0ecec-106">O acesso de representante para suplementos do Outlook está atualmente [em visualização](/office/dev/add-ins/reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview) e é suportado apenas em clientes que são executados no Exchange Online.</span><span class="sxs-lookup"><span data-stu-id="0ecec-106">Delegate access for Outlook add-ins is currently [in preview](/office/dev/add-ins/reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview) and only supported in clients that run against Exchange Online.</span></span> <span data-ttu-id="0ecec-107">Os suplementos que usam esse elemento não podem ser publicados no AppSource nem implantados por meio da implantação centralizada.</span><span class="sxs-lookup"><span data-stu-id="0ecec-107">Add-ins that use this element cannot be published to AppSource or deployed via centralized deployment.</span></span>

<span data-ttu-id="0ecec-108">Veja a seguir um exemplo do elemento **SupportsSharedFolders**.</span><span class="sxs-lookup"><span data-stu-id="0ecec-108">The following is an example of the  **SupportsSharedFolders** element.</span></span>

```XML
...
<VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides" xsi:type="VersionOverridesV1_0">
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides/1.1" xsi:type="VersionOverridesV1_1">
    ...
    <Hosts>
      <Host xsi:type="MailHost">
        <DesktopFormFactor>
          <SupportsSharedFolders>true</SupportsSharedFolders>
          <FunctionFile resid="residDesktopFuncUrl" />
          <ExtensionPoint xsi:type="MessageReadCommandSurface">
            <!-- configure selected extension point -->
          </ExtensionPoint>

          <!-- You can define more than one ExtensionPoint element as needed -->

        </DesktopFormFactor>
      </Host>
    </Hosts>
    ...
  </VersionOverrides>
</VersionOverrides>
...
```
