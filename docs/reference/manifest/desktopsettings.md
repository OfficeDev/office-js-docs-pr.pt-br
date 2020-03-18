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
# <a name="desktopsettings-element"></a><span data-ttu-id="b2662-103">Elemento DesktopSettings</span><span class="sxs-lookup"><span data-stu-id="b2662-103">DesktopSettings element</span></span>

<span data-ttu-id="b2662-104">Especifica o local de origem e as configurações de controle aplicadas quando seu suplemento de email é usado em um computador desktop.</span><span class="sxs-lookup"><span data-stu-id="b2662-104">Specifies source location and control settings that apply when your mail add-in is used on a desktop computer.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="b2662-105">O `DesktopSettings` elemento só está disponível no Outlook clássico na Web (geralmente conectado a versões anteriores do Exchange Server local) e no Outlook 2013 no Windows.</span><span class="sxs-lookup"><span data-stu-id="b2662-105">The `DesktopSettings` element is available only in classic Outlook on the web (usually connected to older versions of on-premises Exchange server) and Outlook 2013 on Windows.</span></span>

<span data-ttu-id="b2662-106">**Tipo de suplemento:** Email</span><span class="sxs-lookup"><span data-stu-id="b2662-106">**Add-in type:** Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="b2662-107">Sintaxe</span><span class="sxs-lookup"><span data-stu-id="b2662-107">Syntax</span></span>

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

## <a name="contained-in"></a><span data-ttu-id="b2662-108">Contido em</span><span class="sxs-lookup"><span data-stu-id="b2662-108">Contained in</span></span>

[<span data-ttu-id="b2662-109">Form</span><span class="sxs-lookup"><span data-stu-id="b2662-109">Form</span></span>](form.md)
