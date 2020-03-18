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
# <a name="form-element"></a><span data-ttu-id="8b0ad-103">Elemento Form</span><span class="sxs-lookup"><span data-stu-id="8b0ad-103">Form element</span></span>

<span data-ttu-id="8b0ad-104">Configurações UX para os formulários que seu suplemento de email usará durante a execução em um determinado dispositivo (área de trabalho, tablet ou telefone).</span><span class="sxs-lookup"><span data-stu-id="8b0ad-104">UX settings for the forms that your mail add-in will use when running on a particular device (desktop, tablet, or phone).</span></span>

> [!IMPORTANT]
> <span data-ttu-id="8b0ad-105">Os `DesktopSettings`elementos `TabletSettings`, e `PhoneSettings` estão disponíveis somente no Outlook clássico na Web (geralmente conectados a versões mais antigas do Exchange Server local) e no Outlook 2013 no Windows.</span><span class="sxs-lookup"><span data-stu-id="8b0ad-105">The `DesktopSettings`, `TabletSettings`, and `PhoneSettings` elements are available only in classic Outlook on the web (usually connected to older versions of on-premises Exchange server) and Outlook 2013 on Windows.</span></span>

<span data-ttu-id="8b0ad-106">**Tipo de suplemento:** Email</span><span class="sxs-lookup"><span data-stu-id="8b0ad-106">**Add-in type:** Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="8b0ad-107">Sintaxe</span><span class="sxs-lookup"><span data-stu-id="8b0ad-107">Syntax</span></span>

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

## <a name="contained-in"></a><span data-ttu-id="8b0ad-108">Contido em</span><span class="sxs-lookup"><span data-stu-id="8b0ad-108">Contained in</span></span>

[<span data-ttu-id="8b0ad-109">FormSettings</span><span class="sxs-lookup"><span data-stu-id="8b0ad-109">FormSettings</span></span>](formsettings.md)


## <a name="can-contain"></a><span data-ttu-id="8b0ad-110">Pode conter</span><span class="sxs-lookup"><span data-stu-id="8b0ad-110">Can contain</span></span>

|<span data-ttu-id="8b0ad-111">**Element**</span><span class="sxs-lookup"><span data-stu-id="8b0ad-111">**Element**</span></span>|
|:-----|
|[<span data-ttu-id="8b0ad-112">DesktopSettings</span><span class="sxs-lookup"><span data-stu-id="8b0ad-112">DesktopSettings</span></span>](desktopsettings.md)|
|[<span data-ttu-id="8b0ad-113">TabletSettings</span><span class="sxs-lookup"><span data-stu-id="8b0ad-113">TabletSettings</span></span>](tabletsettings.md)|
|[<span data-ttu-id="8b0ad-114">PhoneSettings</span><span class="sxs-lookup"><span data-stu-id="8b0ad-114">PhoneSettings</span></span>](phonesettings.md)|
