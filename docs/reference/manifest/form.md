---
title: Elemento Form no arquivo de manifesto
description: Configurações UX para os formulários que seu suplemento de email usará durante a execução em um determinado dispositivo (área de trabalho, tablet ou telefone).
ms.date: 04/09/2020
ms.localizationpriority: medium
ms.openlocfilehash: ecea978cd87a9feed08ba621c9377525acaf4e3a
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 09/12/2021
ms.locfileid: "59148891"
---
# <a name="form-element"></a>Elemento Form

Configurações UX para os formulários que seu suplemento de email usará durante a execução em um determinado dispositivo (área de trabalho, tablet ou telefone).

> [!IMPORTANT]
> Os elementos , e estão disponíveis apenas no Outlook na Web clássico (geralmente conectado a versões mais antigas do servidor Exchange local) e `DesktopSettings` `TabletSettings` Outlook `PhoneSettings` 2013 no Windows.

**Tipo de suplemento:** Email

## <a name="syntax"></a>Sintaxe

```XML
<Form xsi:type="ItemRead">
   <!--https://MyDomain.com/website.html is a placeholder for your own add-in website.-->
   <DesktopSettings>
      <!--If you opt to include RequestedHeight, it must be between 32px to 450px, inclusive.-->
      <RequestedHeight>360</RequestedHeight>
      <SourceLocation DefaultValue="https://MyDomain.com/website.html" />
   </DesktopSettings>
   <TabletSettings>
      <!--If you opt to include RequestedHeight, it must be between 32px to 450px, inclusive.-->
      <RequestedHeight>360</RequestedHeight>
      <SourceLocation DefaultValue="https://MyDomain.com/website.html" />
   </TabletSettings>
   <PhoneSettings>
      <SourceLocation DefaultValue="https://MyDomain.com/website.html" />
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
