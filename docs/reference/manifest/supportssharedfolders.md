---
title: Elemento SupportsSharedFolders no arquivo de manifesto
description: O elemento SupportsSharedFolders define se o suplemento do Outlook está disponível nos cenários de representante.
ms.date: 01/14/2020
localization_priority: Normal
ms.openlocfilehash: 3835f7060cc52a72ff0a5ed4dbdb9f1e09258669
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 06/08/2020
ms.locfileid: "44608709"
---
# <a name="supportssharedfolders-element"></a><span data-ttu-id="8e54d-103">Elemento SupportsSharedFolders</span><span class="sxs-lookup"><span data-stu-id="8e54d-103">SupportsSharedFolders element</span></span>

<span data-ttu-id="8e54d-104">Define se o suplemento do Outlook está disponível nos cenários de representante.</span><span class="sxs-lookup"><span data-stu-id="8e54d-104">Defines whether the Outlook add-in is available in delegate scenarios.</span></span> <span data-ttu-id="8e54d-105">O **SupportsSharedFolders** é um elemento filho de [DesktopFormFactor](desktopformfactor.md).</span><span class="sxs-lookup"><span data-stu-id="8e54d-105">The **SupportsSharedFolders** element is a child element of [DesktopFormFactor](desktopformfactor.md).</span></span> <span data-ttu-id="8e54d-106">Ele é definido como *false* por padrão.</span><span class="sxs-lookup"><span data-stu-id="8e54d-106">It is set to *false* by default.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="8e54d-107">Somente o Outlook na Web e o Windows dão suporte ao elemento **SupportsSharedFolders** .</span><span class="sxs-lookup"><span data-stu-id="8e54d-107">Only Outlook on the web and Windows support the **SupportsSharedFolders** element.</span></span>
>
> <span data-ttu-id="8e54d-108">O suporte para este elemento foi introduzido no conjunto de requisitos 1,8.</span><span class="sxs-lookup"><span data-stu-id="8e54d-108">Support for this element was introduced in requirement set 1.8.</span></span> <span data-ttu-id="8e54d-109">Confira, [clientes e plataformas](../../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) que oferecem suporte a esse conjunto de requisitos.</span><span class="sxs-lookup"><span data-stu-id="8e54d-109">See [clients and platforms](../../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) that support this requirement set.</span></span>

<span data-ttu-id="8e54d-110">Veja a seguir um exemplo do elemento **SupportsSharedFolders** .</span><span class="sxs-lookup"><span data-stu-id="8e54d-110">The following is an example of the **SupportsSharedFolders** element.</span></span>

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
