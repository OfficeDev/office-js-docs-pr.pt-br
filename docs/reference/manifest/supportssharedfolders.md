---
title: Elemento SupportsSharedFolders no arquivo de manifesto
description: O elemento SupportsSharedFolders define se o Outlook está disponível em pastas compartilhadas e cenários de caixa de correio compartilhadas.
ms.date: 06/15/2021
ms.localizationpriority: medium
ms.openlocfilehash: f6a7cbe1c6549c8a93a6ecceab9e4bdaba07001f
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 09/12/2021
ms.locfileid: "59151993"
---
# <a name="supportssharedfolders-element"></a>Elemento SupportsSharedFolders

Define se o Outlook está disponível em cenários de caixa de correio compartilhada (agora em visualização) e pastas compartilhadas (ou seja, acesso de representante). O **SupportsSharedFolders** é um elemento filho de [DesktopFormFactor](desktopformfactor.md). Ele é definido como *false* por padrão.

> [!IMPORTANT]
> O suporte a esse elemento foi introduzido no conjunto de requisitos 1.8. Confira, [clientes e plataformas](../../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) que oferecem suporte a esse conjunto de requisitos.

A seguir, um exemplo do **elemento SupportsSharedFolders.**

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
