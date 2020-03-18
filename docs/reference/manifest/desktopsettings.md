---
title: Elemento DesktopSettings no arquivo de manifesto
description: Especifica o local de origem e as configurações de controle aplicadas quando seu suplemento de email é usado em um computador desktop.
ms.date: 01/14/2020
localization_priority: Normal
ms.openlocfilehash: d48532482fc71fec2a96133ee8e813cae798613f
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/17/2020
ms.locfileid: "42718353"
---
# <a name="desktopsettings-element"></a>Elemento DesktopSettings

Especifica o local de origem e as configurações de controle aplicadas quando seu suplemento de email é usado em um computador desktop.

> [!IMPORTANT]
> O `DesktopSettings` elemento só está disponível no Outlook clássico na Web (geralmente conectado a versões anteriores do Exchange Server local) e no Outlook 2013 no Windows.

**Tipo de suplemento:** Email

## <a name="syntax"></a>Sintaxe

```XML
<Form xsi:type="ItemRead">
   <!--website.html is a placeholder for your own add-in website.-->
   <DesktopSettings>
      <SourceLocation DefaultValue="https://website.html" />
      <!--RequestedHeight must be between 240px to 800px, inclusive.-->
      <RequestedHeight>360</RequestedHeight>
   </DesktopSettings>
   <TabletSettings>
      <SourceLocation DefaultValue="https://website.html" />
      <!--RequestedHeight must be between 240px to 800px, inclusive.-->
      <RequestedHeight>360</RequestedHeight>
   </TabletSettings>
   <PhoneSettings>
      <SourceLocation DefaultValue="https://website.html" />
   </PhoneSettings>
</Form>
```

## <a name="contained-in"></a>Contido em

[Form](form.md)
