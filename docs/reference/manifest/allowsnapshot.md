---
title: Elemento AllowSnapshot no arquivo de manifesto
description: Especifica se o instantâneo de uma imagem do suplemento de conteúdo é salvo com o documento host.
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: c46dcd882592c0b015dae4b9774533b96fe75cfe
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 06/08/2020
ms.locfileid: "44608786"
---
# <a name="allowsnapshot-element"></a><span data-ttu-id="15c2e-103">Elemento AllowSnapshot</span><span class="sxs-lookup"><span data-stu-id="15c2e-103">AllowSnapshot element</span></span>

<span data-ttu-id="15c2e-104">Especifica se o instantâneo de uma imagem do suplemento de conteúdo é salvo com o documento host.</span><span class="sxs-lookup"><span data-stu-id="15c2e-104">Specifies whether a snapshot image of your content add-in is saved with the host document.</span></span>

<span data-ttu-id="15c2e-105">**Tipo de suplemento:** Conteúdo</span><span class="sxs-lookup"><span data-stu-id="15c2e-105">**Add-in type:** Content</span></span>

## <a name="syntax"></a><span data-ttu-id="15c2e-106">Sintaxe</span><span class="sxs-lookup"><span data-stu-id="15c2e-106">Syntax</span></span>

```XML
<AllowSnapshot> [true | false]</AllowSnapshot>
```

## <a name="contained-in"></a><span data-ttu-id="15c2e-107">Contido em</span><span class="sxs-lookup"><span data-stu-id="15c2e-107">Contained in</span></span>

[<span data-ttu-id="15c2e-108">OfficeApp</span><span class="sxs-lookup"><span data-stu-id="15c2e-108">OfficeApp</span></span>](officeapp.md)

## <a name="remarks"></a><span data-ttu-id="15c2e-109">Comentários</span><span class="sxs-lookup"><span data-stu-id="15c2e-109">Remarks</span></span>

 > [!IMPORTANT]
 > <span data-ttu-id="15c2e-110">**AllowSnapshot** é `true` por padrão.</span><span class="sxs-lookup"><span data-stu-id="15c2e-110">**AllowSnapshot** is `true` by default.</span></span> <span data-ttu-id="15c2e-111">Isso cria uma imagem do suplemento visível para os usuários que abrirem o documento em uma versão do aplicativo host que não oferece suporte a Suplementos do Office,ou fornece uma imagem estática do suplemento se o aplicativo host não se conectar ao servidor que hospeda o suplemento.</span><span class="sxs-lookup"><span data-stu-id="15c2e-111">This makes an image of the add-in visible for users that open the document in a version of the host application that doesn't support Office Add-ins, or provides a static image of the add-in if the host application can't connect to the server hosting the add-in.</span></span> <span data-ttu-id="15c2e-112">However, this also means that potentially sensitive information displayed in the add-in can be accessed directly from the document hosting the add-in.</span><span class="sxs-lookup"><span data-stu-id="15c2e-112">However, this also means that potentially sensitive information displayed in the add-in can be accessed directly from the document hosting the add-in.</span></span>

