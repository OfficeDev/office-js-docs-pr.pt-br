---
title: Elemento Form no arquivo de manifesto
description: Configurações UX para os formulários que seu suplemento de email usará durante a execução em um determinado dispositivo (área de trabalho, tablet ou telefone).
ms.date: 04/09/2020
localization_priority: Normal
ms.openlocfilehash: 3e8d60c13a72a50090075d7cd16a0719498c4982
ms.sourcegitcommit: c6e3bfd3deb77982d0b7082afd6a48678e96e1c3
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 04/10/2020
ms.locfileid: "43215065"
---
# <a name="form-element"></a><span data-ttu-id="64b29-103">Elemento Form</span><span class="sxs-lookup"><span data-stu-id="64b29-103">Form element</span></span>

<span data-ttu-id="64b29-104">Configurações UX para os formulários que seu suplemento de email usará durante a execução em um determinado dispositivo (área de trabalho, tablet ou telefone).</span><span class="sxs-lookup"><span data-stu-id="64b29-104">UX settings for the forms that your mail add-in will use when running on a particular device (desktop, tablet, or phone).</span></span>

> [!IMPORTANT]
> <span data-ttu-id="64b29-105">Os `DesktopSettings`elementos `TabletSettings`, e `PhoneSettings` estão disponíveis somente no Outlook clássico na Web (geralmente conectados a versões mais antigas do Exchange Server local) e no Outlook 2013 no Windows.</span><span class="sxs-lookup"><span data-stu-id="64b29-105">The `DesktopSettings`, `TabletSettings`, and `PhoneSettings` elements are available only in classic Outlook on the web (usually connected to older versions of on-premises Exchange server) and Outlook 2013 on Windows.</span></span>

<span data-ttu-id="64b29-106">**Tipo de suplemento:** Email</span><span class="sxs-lookup"><span data-stu-id="64b29-106">**Add-in type:** Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="64b29-107">Sintaxe</span><span class="sxs-lookup"><span data-stu-id="64b29-107">Syntax</span></span>

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

## <a name="contained-in"></a><span data-ttu-id="64b29-108">Contido em</span><span class="sxs-lookup"><span data-stu-id="64b29-108">Contained in</span></span>

[<span data-ttu-id="64b29-109">FormSettings</span><span class="sxs-lookup"><span data-stu-id="64b29-109">FormSettings</span></span>](formsettings.md)


## <a name="can-contain"></a><span data-ttu-id="64b29-110">Pode conter</span><span class="sxs-lookup"><span data-stu-id="64b29-110">Can contain</span></span>

|<span data-ttu-id="64b29-111">**Element**</span><span class="sxs-lookup"><span data-stu-id="64b29-111">**Element**</span></span>|
|:-----|
|[<span data-ttu-id="64b29-112">DesktopSettings</span><span class="sxs-lookup"><span data-stu-id="64b29-112">DesktopSettings</span></span>](desktopsettings.md)|
|[<span data-ttu-id="64b29-113">TabletSettings</span><span class="sxs-lookup"><span data-stu-id="64b29-113">TabletSettings</span></span>](tabletsettings.md)|
|[<span data-ttu-id="64b29-114">PhoneSettings</span><span class="sxs-lookup"><span data-stu-id="64b29-114">PhoneSettings</span></span>](phonesettings.md)|
