---
title: Elemento TabletSettings no arquivo de manifesto
description: ''
ms.date: 01/13/2020
localization_priority: Normal
ms.openlocfilehash: 977fc2a781f3b93e4eb36041473c683196314adb
ms.sourcegitcommit: dc42e0276007f8ab006028b9cd0cc1526c1bd100
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 01/14/2020
ms.locfileid: "41120617"
---
# <a name="tabletsettings-element"></a><span data-ttu-id="8c23f-102">Elemento TabletSettings</span><span class="sxs-lookup"><span data-stu-id="8c23f-102">TabletSettings element</span></span>

<span data-ttu-id="8c23f-103">Especifica as configurações de controle aplicadas quando seu suplemento de email é usado em um tablet.</span><span class="sxs-lookup"><span data-stu-id="8c23f-103">Specifies control settings that apply when your mail add-in is used on a tablet.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="8c23f-104">O `TabletSettings` elemento só está disponível no Outlook clássico na Web (geralmente conectado a versões anteriores do Exchange Server local) e no Outlook 2013 no Windows.</span><span class="sxs-lookup"><span data-stu-id="8c23f-104">The `TabletSettings` element is available only in classic Outlook on the web (usually connected to older versions of on-premises Exchange server) and Outlook 2013 on Windows.</span></span> <span data-ttu-id="8c23f-105">Para dar suporte ao Outlook no Android e iOS, confira [suplementos do Outlook Mobile](/outlook/add-ins/outlook-mobile-addins).</span><span class="sxs-lookup"><span data-stu-id="8c23f-105">To support Outlook on Android and iOS, see [Add-ins for Outlook Mobile](/outlook/add-ins/outlook-mobile-addins).</span></span>

<span data-ttu-id="8c23f-106">**Tipo de suplemento:** Email</span><span class="sxs-lookup"><span data-stu-id="8c23f-106">**Add-in type:** Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="8c23f-107">Sintaxe</span><span class="sxs-lookup"><span data-stu-id="8c23f-107">Syntax</span></span>

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

## <a name="contained-in"></a><span data-ttu-id="8c23f-108">Contido em</span><span class="sxs-lookup"><span data-stu-id="8c23f-108">Contained in</span></span>

[<span data-ttu-id="8c23f-109">Form</span><span class="sxs-lookup"><span data-stu-id="8c23f-109">Form</span></span>](form.md)

