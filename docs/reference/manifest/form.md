---
title: Elemento Form no arquivo de manifesto
description: Configurações UX para os formulários que seu suplemento de email usará durante a execução em um determinado dispositivo (área de trabalho, tablet ou telefone).
ms.date: 01/13/2020
localization_priority: Normal
ms.openlocfilehash: 9b1696b2fecf6b07ee2a3c0a31611d4f2ad1f291
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/17/2020
ms.locfileid: "42718206"
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
