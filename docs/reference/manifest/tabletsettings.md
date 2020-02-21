---
title: Elemento TabletSettings no arquivo de manifesto
description: ''
ms.date: 01/13/2020
localization_priority: Normal
ms.openlocfilehash: bf11dcfec4dfe2c40764722d23c7a69c289bba65
ms.sourcegitcommit: a3ddfdb8a95477850148c4177e20e56a8673517c
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 02/20/2020
ms.locfileid: "42163861"
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

