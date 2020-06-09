---
title: Elemento TabletSettings no arquivo de manifesto
description: O elemento TabletSettings especifica as configurações de controle que se aplicam quando seu suplemento de email é usado em um Tablet.
ms.date: 04/09/2020
localization_priority: Normal
ms.openlocfilehash: b5a74db4f9fb43df10a08ab43b59507f6e0d7952
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 06/08/2020
ms.locfileid: "44608695"
---
# <a name="tabletsettings-element"></a><span data-ttu-id="da3d2-103">Elemento TabletSettings</span><span class="sxs-lookup"><span data-stu-id="da3d2-103">TabletSettings element</span></span>

<span data-ttu-id="da3d2-104">Especifica as configurações de controle aplicadas quando seu suplemento de email é usado em um tablet.</span><span class="sxs-lookup"><span data-stu-id="da3d2-104">Specifies control settings that apply when your mail add-in is used on a tablet.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="da3d2-105">O `TabletSettings` elemento só está disponível no Outlook clássico na Web (geralmente conectado a versões anteriores do Exchange Server local) e no outlook 2013 no Windows.</span><span class="sxs-lookup"><span data-stu-id="da3d2-105">The `TabletSettings` element is available only in classic Outlook on the web (usually connected to older versions of on-premises Exchange server) and Outlook 2013 on Windows.</span></span> <span data-ttu-id="da3d2-106">Para dar suporte ao Outlook no Android e iOS, confira [suplementos do Outlook Mobile](../../outlook/outlook-mobile-addins.md).</span><span class="sxs-lookup"><span data-stu-id="da3d2-106">To support Outlook on Android and iOS, see [Add-ins for Outlook Mobile](../../outlook/outlook-mobile-addins.md).</span></span>

<span data-ttu-id="da3d2-107">**Tipo de suplemento:** Email</span><span class="sxs-lookup"><span data-stu-id="da3d2-107">**Add-in type:** Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="da3d2-108">Sintaxe</span><span class="sxs-lookup"><span data-stu-id="da3d2-108">Syntax</span></span>

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

## <a name="contained-in"></a><span data-ttu-id="da3d2-109">Contido em</span><span class="sxs-lookup"><span data-stu-id="da3d2-109">Contained in</span></span>

[<span data-ttu-id="da3d2-110">Form</span><span class="sxs-lookup"><span data-stu-id="da3d2-110">Form</span></span>](form.md)
