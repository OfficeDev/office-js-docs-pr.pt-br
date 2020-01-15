---
title: Elemento Form no arquivo de manifesto
description: ''
ms.date: 01/13/2020
localization_priority: Normal
ms.openlocfilehash: d545d471e007f0077a8310b0b847bbbf99a8f7ac
ms.sourcegitcommit: dc42e0276007f8ab006028b9cd0cc1526c1bd100
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 01/14/2020
ms.locfileid: "41120646"
---
# <a name="form-element"></a>Elemento Form

Configurações UX para os formulários que seu suplemento de email usará durante a execução em um determinado dispositivo (área de trabalho, tablet ou telefone).

> [!IMPORTANT]
> Os `DesktopSettings`elementos `TabletSettings`, e `PhoneSettings` estão disponíveis somente no Outlook clássico na Web (geralmente conectados a versões mais antigas do Exchange Server local) e no Outlook 2013 no Windows.

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

[FormSettings](formsettings.md)


## <a name="can-contain"></a>Pode conter

|**Element**|
|:-----|
|[DesktopSettings](desktopsettings.md)|
|[TabletSettings](tabletsettings.md)|
|[PhoneSettings](phonesettings.md)|
