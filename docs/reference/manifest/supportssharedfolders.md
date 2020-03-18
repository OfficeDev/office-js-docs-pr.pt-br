---
title: Elemento SupportsSharedFolders no arquivo de manifesto
description: O elemento SupportsSharedFolders define se o suplemento do Outlook está disponível nos cenários de representante.
ms.date: 01/14/2020
localization_priority: Normal
ms.openlocfilehash: 66a426b0c31bda61feb23cb83d63722898dfb503
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/17/2020
ms.locfileid: "42717884"
---
# <a name="supportssharedfolders-element"></a>Elemento SupportsSharedFolders

Define se o suplemento do Outlook está disponível nos cenários de representante. O **SupportsSharedFolders** é um elemento filho de [DesktopFormFactor](desktopformfactor.md). Ele é definido como *false* por padrão.

> [!IMPORTANT]
> Somente o Outlook na Web e o Windows dão suporte ao elemento **SupportsSharedFolders** .
>
> O suporte para este elemento foi introduzido no conjunto de requisitos 1,8. Confira, [clientes e plataformas](../../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) que oferecem suporte a esse conjunto de requisitos.

Veja a seguir um exemplo do elemento **SupportsSharedFolders** .

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
