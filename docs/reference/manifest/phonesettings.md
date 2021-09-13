---
title: Elemento PhoneSettings no arquivo de manifesto
description: O elemento PhoneSettings especifica o local de origem e as configurações de controle que se aplicam quando o seu complemento de email é usado em um telefone.
ms.date: 04/09/2020
ms.localizationpriority: medium
ms.openlocfilehash: 1e52827a20ee95397541f7c1d54c732ff8f96ba5
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 09/12/2021
ms.locfileid: "59152029"
---
# <a name="phonesettings-element"></a>Elemento PhoneSettings

Especifica o local de origem e as configurações de controle aplicadas quando o seu suplemento de email é usado em um telefone.

> [!IMPORTANT]
> O elemento está disponível apenas no Outlook na Web clássico (geralmente conectado a versões mais antigas do servidor Exchange local) e Outlook `PhoneSettings` 2013 no Windows. Para dar Outlook em Android e iOS, consulte [Add-ins for Outlook Mobile](../../outlook/outlook-mobile-addins.md).

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

[Form](form.md)

