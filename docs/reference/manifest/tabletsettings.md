---
title: Elemento TabletSettings no arquivo de manifesto
description: O elemento TabletSettings especifica as configurações de controle que se aplicam quando seu suplemento de email é usado em um Tablet.
ms.date: 01/13/2020
localization_priority: Normal
ms.openlocfilehash: 2b8b372d27274d89d3aed4b5bacb9faa4893fda5
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/17/2020
ms.locfileid: "42717856"
---
# <a name="tabletsettings-element"></a>Elemento TabletSettings

Especifica as configurações de controle aplicadas quando seu suplemento de email é usado em um tablet.

> [!IMPORTANT]
> O `TabletSettings` elemento só está disponível no Outlook clássico na Web (geralmente conectado a versões anteriores do Exchange Server local) e no Outlook 2013 no Windows. Para dar suporte ao Outlook no Android e iOS, confira [suplementos do Outlook Mobile](../../outlook/outlook-mobile-addins.md).

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

