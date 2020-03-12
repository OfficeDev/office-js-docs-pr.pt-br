---
title: Elemento SupportsSharedFolders no arquivo de manifesto
description: ''
ms.date: 01/14/2020
localization_priority: Normal
ms.openlocfilehash: 81401b79f4c443305e376df7a66a07d916393d17
ms.sourcegitcommit: 4079903c3cc45b7d8c041509a44e9fc38da399b1
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/11/2020
ms.locfileid: "42596750"
---
# <a name="supportssharedfolders-element"></a><span data-ttu-id="dc932-102">Elemento SupportsSharedFolders</span><span class="sxs-lookup"><span data-stu-id="dc932-102">SupportsSharedFolders element</span></span>

<span data-ttu-id="dc932-103">Define se o suplemento do Outlook está disponível nos cenários de representante.</span><span class="sxs-lookup"><span data-stu-id="dc932-103">Defines whether the Outlook add-in is available in delegate scenarios.</span></span> <span data-ttu-id="dc932-104">O **SupportsSharedFolders** é um elemento filho de [DesktopFormFactor](desktopformfactor.md).</span><span class="sxs-lookup"><span data-stu-id="dc932-104">The **SupportsSharedFolders** element is a child element of [DesktopFormFactor](desktopformfactor.md).</span></span> <span data-ttu-id="dc932-105">Ele é definido como *false* por padrão.</span><span class="sxs-lookup"><span data-stu-id="dc932-105">It is set to *false* by default.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="dc932-106">Somente o Outlook na Web e o Windows dão suporte ao elemento **SupportsSharedFolders** .</span><span class="sxs-lookup"><span data-stu-id="dc932-106">Only Outlook on the web and Windows support the **SupportsSharedFolders** element.</span></span>
>
> <span data-ttu-id="dc932-107">O suporte para este elemento foi introduzido no conjunto de requisitos 1,8.</span><span class="sxs-lookup"><span data-stu-id="dc932-107">Support for this element was introduced in requirement set 1.8.</span></span> <span data-ttu-id="dc932-108">Confira, [clientes e plataformas](../../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) que oferecem suporte a esse conjunto de requisitos.</span><span class="sxs-lookup"><span data-stu-id="dc932-108">See [clients and platforms](../../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) that support this requirement set.</span></span>

<span data-ttu-id="dc932-109">Veja a seguir um exemplo do elemento **SupportsSharedFolders** .</span><span class="sxs-lookup"><span data-stu-id="dc932-109">The following is an example of the **SupportsSharedFolders** element.</span></span>

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
            <!-- Configure selected extension point. -->
          </ExtensionPoint>

          <!-- You can define more than one ExtensionPoint element as needed. -->

        </DesktopFormFactor>
      </Host>
    </Hosts>
    ...
  </VersionOverrides>
</VersionOverrides>
...
```
