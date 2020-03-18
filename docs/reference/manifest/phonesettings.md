---
title: Elemento PhoneSettings no arquivo de manifesto
description: O elemento PhoneSettings especifica o local de origem e as configurações de controle que se aplicam quando seu suplemento de email é usado em um telefone.
ms.date: 01/13/2020
localization_priority: Normal
ms.openlocfilehash: 581a3ae71a58cd05aac52129a6f4395a60c20cef
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/17/2020
ms.locfileid: "42720471"
---
# <a name="phonesettings-element"></a><span data-ttu-id="1a8ed-103">Elemento PhoneSettings</span><span class="sxs-lookup"><span data-stu-id="1a8ed-103">PhoneSettings element</span></span>

<span data-ttu-id="1a8ed-104">Especifica o local de origem e as configurações de controle aplicadas quando o seu suplemento de email é usado em um telefone.</span><span class="sxs-lookup"><span data-stu-id="1a8ed-104">Specifies source location and control settings that apply when your mail add-in is used on a phone.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="1a8ed-105">O `PhoneSettings` elemento só está disponível no Outlook clássico na Web (geralmente conectado a versões anteriores do Exchange Server local) e no Outlook 2013 no Windows.</span><span class="sxs-lookup"><span data-stu-id="1a8ed-105">The `PhoneSettings` element is available only in classic Outlook on the web (usually connected to older versions of on-premises Exchange server) and Outlook 2013 on Windows.</span></span> <span data-ttu-id="1a8ed-106">Para dar suporte ao Outlook no Android e iOS, confira [suplementos do Outlook Mobile](../../outlook/outlook-mobile-addins.md).</span><span class="sxs-lookup"><span data-stu-id="1a8ed-106">To support Outlook on Android and iOS, see [Add-ins for Outlook Mobile](../../outlook/outlook-mobile-addins.md).</span></span>

<span data-ttu-id="1a8ed-107">**Tipo de suplemento:** Email</span><span class="sxs-lookup"><span data-stu-id="1a8ed-107">**Add-in type:** Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="1a8ed-108">Sintaxe</span><span class="sxs-lookup"><span data-stu-id="1a8ed-108">Syntax</span></span>

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

## <a name="contained-in"></a><span data-ttu-id="1a8ed-109">Contido em</span><span class="sxs-lookup"><span data-stu-id="1a8ed-109">Contained in</span></span>

[<span data-ttu-id="1a8ed-110">Form</span><span class="sxs-lookup"><span data-stu-id="1a8ed-110">Form</span></span>](form.md)

