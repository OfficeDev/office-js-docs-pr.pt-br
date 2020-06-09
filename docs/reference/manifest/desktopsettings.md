---
title: Elemento DesktopSettings no arquivo de manifesto
description: Especifica o local de origem e as configurações de controle aplicadas quando seu suplemento de email é usado em um computador desktop.
ms.date: 04/09/2020
localization_priority: Normal
ms.openlocfilehash: 50201080d8be3c8943d16730c34a4bac236d7b90
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 06/08/2020
ms.locfileid: "44612273"
---
# <a name="desktopsettings-element"></a><span data-ttu-id="32c48-103">Elemento DesktopSettings</span><span class="sxs-lookup"><span data-stu-id="32c48-103">DesktopSettings element</span></span>

<span data-ttu-id="32c48-104">Especifica o local de origem e as configurações de controle aplicadas quando seu suplemento de email é usado em um computador desktop.</span><span class="sxs-lookup"><span data-stu-id="32c48-104">Specifies source location and control settings that apply when your mail add-in is used on a desktop computer.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="32c48-105">O `DesktopSettings` elemento só está disponível no Outlook clássico na Web (geralmente conectado a versões anteriores do Exchange Server local) e no outlook 2013 no Windows.</span><span class="sxs-lookup"><span data-stu-id="32c48-105">The `DesktopSettings` element is available only in classic Outlook on the web (usually connected to older versions of on-premises Exchange server) and Outlook 2013 on Windows.</span></span>

<span data-ttu-id="32c48-106">**Tipo de suplemento:** Email</span><span class="sxs-lookup"><span data-stu-id="32c48-106">**Add-in type:** Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="32c48-107">Sintaxe</span><span class="sxs-lookup"><span data-stu-id="32c48-107">Syntax</span></span>

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

## <a name="contained-in"></a><span data-ttu-id="32c48-108">Contido em</span><span class="sxs-lookup"><span data-stu-id="32c48-108">Contained in</span></span>

[<span data-ttu-id="32c48-109">Form</span><span class="sxs-lookup"><span data-stu-id="32c48-109">Form</span></span>](form.md)
