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
# <a name="form-element"></a><span data-ttu-id="d9d06-102">Elemento Form</span><span class="sxs-lookup"><span data-stu-id="d9d06-102">Form element</span></span>

<span data-ttu-id="d9d06-103">Configurações UX para os formulários que seu suplemento de email usará durante a execução em um determinado dispositivo (área de trabalho, tablet ou telefone).</span><span class="sxs-lookup"><span data-stu-id="d9d06-103">UX settings for the forms that your mail add-in will use when running on a particular device (desktop, tablet, or phone).</span></span>

> [!IMPORTANT]
> <span data-ttu-id="d9d06-104">Os `DesktopSettings`elementos `TabletSettings`, e `PhoneSettings` estão disponíveis somente no Outlook clássico na Web (geralmente conectados a versões mais antigas do Exchange Server local) e no Outlook 2013 no Windows.</span><span class="sxs-lookup"><span data-stu-id="d9d06-104">The `DesktopSettings`, `TabletSettings`, and `PhoneSettings` elements are available only in classic Outlook on the web (usually connected to older versions of on-premises Exchange server) and Outlook 2013 on Windows.</span></span>

<span data-ttu-id="d9d06-105">**Tipo de suplemento:** Email</span><span class="sxs-lookup"><span data-stu-id="d9d06-105">**Add-in type:** Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="d9d06-106">Sintaxe</span><span class="sxs-lookup"><span data-stu-id="d9d06-106">Syntax</span></span>

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

## <a name="contained-in"></a><span data-ttu-id="d9d06-107">Contido em</span><span class="sxs-lookup"><span data-stu-id="d9d06-107">Contained in</span></span>

[<span data-ttu-id="d9d06-108">FormSettings</span><span class="sxs-lookup"><span data-stu-id="d9d06-108">FormSettings</span></span>](formsettings.md)


## <a name="can-contain"></a><span data-ttu-id="d9d06-109">Pode conter</span><span class="sxs-lookup"><span data-stu-id="d9d06-109">Can contain</span></span>

|<span data-ttu-id="d9d06-110">**Element**</span><span class="sxs-lookup"><span data-stu-id="d9d06-110">**Element**</span></span>|
|:-----|
|[<span data-ttu-id="d9d06-111">DesktopSettings</span><span class="sxs-lookup"><span data-stu-id="d9d06-111">DesktopSettings</span></span>](desktopsettings.md)|
|[<span data-ttu-id="d9d06-112">TabletSettings</span><span class="sxs-lookup"><span data-stu-id="d9d06-112">TabletSettings</span></span>](tabletsettings.md)|
|[<span data-ttu-id="d9d06-113">PhoneSettings</span><span class="sxs-lookup"><span data-stu-id="d9d06-113">PhoneSettings</span></span>](phonesettings.md)|
