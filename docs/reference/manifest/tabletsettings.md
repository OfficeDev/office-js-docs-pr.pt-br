---
title: Elemento TabletSettings no arquivo de manifesto
description: ''
ms.date: 01/13/2020
localization_priority: Normal
ms.openlocfilehash: 977fc2a781f3b93e4eb36041473c683196314adb
ms.sourcegitcommit: dc42e0276007f8ab006028b9cd0cc1526c1bd100
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 01/14/2020
ms.locfileid: "41120617"
---
# <a name="tabletsettings-element"></a>Elemento TabletSettings

Especifica as configurações de controle aplicadas quando seu suplemento de email é usado em um tablet.

> [!IMPORTANT]
> O `TabletSettings` elemento só está disponível no Outlook clássico na Web (geralmente conectado a versões anteriores do Exchange Server local) e no Outlook 2013 no Windows. Para dar suporte ao Outlook no Android e iOS, confira [suplementos do Outlook Mobile](/outlook/add-ins/outlook-mobile-addins).

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

