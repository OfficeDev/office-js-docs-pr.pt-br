---
title: Elemento DesktopSettings no arquivo de manifesto
description: ''
ms.date: 01/14/2020
localization_priority: Normal
ms.openlocfilehash: 6dfa69d407e267a1cbcfdeaad0bdf9cdf75c1465
ms.sourcegitcommit: dc42e0276007f8ab006028b9cd0cc1526c1bd100
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 01/14/2020
ms.locfileid: "41120639"
---
# <a name="desktopsettings-element"></a><span data-ttu-id="1b08b-102">Elemento DesktopSettings</span><span class="sxs-lookup"><span data-stu-id="1b08b-102">DesktopSettings element</span></span>

<span data-ttu-id="1b08b-103">Especifica o local de origem e as configurações de controle aplicadas quando seu suplemento de email é usado em um computador desktop.</span><span class="sxs-lookup"><span data-stu-id="1b08b-103">Specifies source location and control settings that apply when your mail add-in is used on a desktop computer.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="1b08b-104">O `DesktopSettings` elemento só está disponível no Outlook clássico na Web (geralmente conectado a versões anteriores do Exchange Server local) e no Outlook 2013 no Windows.</span><span class="sxs-lookup"><span data-stu-id="1b08b-104">The `DesktopSettings` element is available only in classic Outlook on the web (usually connected to older versions of on-premises Exchange server) and Outlook 2013 on Windows.</span></span>

<span data-ttu-id="1b08b-105">**Tipo de suplemento:** Email</span><span class="sxs-lookup"><span data-stu-id="1b08b-105">**Add-in type:** Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="1b08b-106">Sintaxe</span><span class="sxs-lookup"><span data-stu-id="1b08b-106">Syntax</span></span>

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

## <a name="contained-in"></a><span data-ttu-id="1b08b-107">Contido em</span><span class="sxs-lookup"><span data-stu-id="1b08b-107">Contained in</span></span>

[<span data-ttu-id="1b08b-108">Form</span><span class="sxs-lookup"><span data-stu-id="1b08b-108">Form</span></span>](form.md)
