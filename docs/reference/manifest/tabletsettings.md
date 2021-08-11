---
title: Elemento TabletSettings no arquivo de manifesto
description: O elemento TabletSettings especifica as configurações de controle que se aplicam quando o seu complemento de email é usado em um tablet.
ms.date: 04/09/2020
localization_priority: Normal
ms.openlocfilehash: 3360a446396da058b5ced0127050d807f33e7bda007b1da88414b351782d0127
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 08/07/2021
ms.locfileid: "57094639"
---
# <a name="tabletsettings-element"></a>Elemento TabletSettings

Especifica as configurações de controle aplicadas quando seu suplemento de email é usado em um tablet.

> [!IMPORTANT]
> O elemento está disponível apenas no Outlook na Web clássico (geralmente conectado a versões mais antigas do servidor Exchange local) e Outlook `TabletSettings` 2013 no Windows. Para dar Outlook em Android e iOS, consulte [Add-ins for Outlook Mobile](../../outlook/outlook-mobile-addins.md).

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
