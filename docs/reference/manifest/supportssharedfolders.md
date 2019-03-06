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
# <a name="supportssharedfolders-element"></a>Elemento SupportsSharedFolders

Define se o suplemento do Outlook está disponível nos cenários de representante. O **SupportsSharedFolders** é um elemento filho de [DesktopFormFactor](desktopformfactor.md). Ele é definido como *false* por padrão.

> [!IMPORTANT]
> O acesso de representante para suplementos do Outlook está atualmente em visualização e é suportado apenas em clientes que são executados no Exchange Online. Os suplementos que usam esse elemento não podem ser publicados no AppSource nem implantados por meio da implantação centralizada.

Veja a seguir um exemplo do elemento **SupportsSharedFolders**.

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
