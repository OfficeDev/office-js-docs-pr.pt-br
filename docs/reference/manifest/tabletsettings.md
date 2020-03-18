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
# <a name="tabletsettings-element"></a><span data-ttu-id="72624-103">Elemento TabletSettings</span><span class="sxs-lookup"><span data-stu-id="72624-103">TabletSettings element</span></span>

<span data-ttu-id="72624-104">Especifica as configurações de controle aplicadas quando seu suplemento de email é usado em um tablet.</span><span class="sxs-lookup"><span data-stu-id="72624-104">Specifies control settings that apply when your mail add-in is used on a tablet.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="72624-105">O `TabletSettings` elemento só está disponível no Outlook clássico na Web (geralmente conectado a versões anteriores do Exchange Server local) e no Outlook 2013 no Windows.</span><span class="sxs-lookup"><span data-stu-id="72624-105">The `TabletSettings` element is available only in classic Outlook on the web (usually connected to older versions of on-premises Exchange server) and Outlook 2013 on Windows.</span></span> <span data-ttu-id="72624-106">Para dar suporte ao Outlook no Android e iOS, confira [suplementos do Outlook Mobile](../../outlook/outlook-mobile-addins.md).</span><span class="sxs-lookup"><span data-stu-id="72624-106">To support Outlook on Android and iOS, see [Add-ins for Outlook Mobile](../../outlook/outlook-mobile-addins.md).</span></span>

<span data-ttu-id="72624-107">**Tipo de suplemento:** Email</span><span class="sxs-lookup"><span data-stu-id="72624-107">**Add-in type:** Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="72624-108">Sintaxe</span><span class="sxs-lookup"><span data-stu-id="72624-108">Syntax</span></span>

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

## <a name="contained-in"></a><span data-ttu-id="72624-109">Contido em</span><span class="sxs-lookup"><span data-stu-id="72624-109">Contained in</span></span>

[<span data-ttu-id="72624-110">Form</span><span class="sxs-lookup"><span data-stu-id="72624-110">Form</span></span>](form.md)

