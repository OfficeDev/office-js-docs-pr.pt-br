---
title: Elemento DesktopSettings no arquivo de manifesto
description: Especifica o local de origem e as configurações de controle aplicadas quando seu suplemento de email é usado em um computador desktop.
ms.date: 04/09/2020
localization_priority: Normal
ms.openlocfilehash: d8f35a29e14337a849f81b0becb60f761116cb48fa420119228255bb1179bb35
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 08/07/2021
ms.locfileid: "57089846"
---
# <a name="desktopsettings-element"></a>Elemento DesktopSettings

Especifica o local de origem e as configurações de controle aplicadas quando seu suplemento de email é usado em um computador desktop.

> [!IMPORTANT]
> O elemento está disponível apenas no Outlook na Web clássico (geralmente conectado a versões mais antigas do servidor Exchange local) e Outlook `DesktopSettings` 2013 no Windows.

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
