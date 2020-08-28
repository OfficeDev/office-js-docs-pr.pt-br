---
title: Elemento AllowSnapshot no arquivo de manifesto
description: Especifica se o instantâneo de uma imagem do suplemento de conteúdo é salvo com o documento host.
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: ea910e1ad747e304dbc6ab4fbdcf44a9610dab19
ms.sourcegitcommit: 9609bd5b4982cdaa2ea7637709a78a45835ffb19
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 08/28/2020
ms.locfileid: "47294273"
---
# <a name="allowsnapshot-element"></a><span data-ttu-id="275de-103">Elemento AllowSnapshot</span><span class="sxs-lookup"><span data-stu-id="275de-103">AllowSnapshot element</span></span>

<span data-ttu-id="275de-104">Especifica se o instantâneo de uma imagem do suplemento de conteúdo é salvo com o documento host.</span><span class="sxs-lookup"><span data-stu-id="275de-104">Specifies whether a snapshot image of your content add-in is saved with the host document.</span></span>

<span data-ttu-id="275de-105">**Tipo de suplemento:** Conteúdo</span><span class="sxs-lookup"><span data-stu-id="275de-105">**Add-in type:** Content</span></span>

## <a name="syntax"></a><span data-ttu-id="275de-106">Sintaxe</span><span class="sxs-lookup"><span data-stu-id="275de-106">Syntax</span></span>

```XML
<AllowSnapshot> [true | false]</AllowSnapshot>
```

## <a name="contained-in"></a><span data-ttu-id="275de-107">Contido em</span><span class="sxs-lookup"><span data-stu-id="275de-107">Contained in</span></span>

[<span data-ttu-id="275de-108">OfficeApp</span><span class="sxs-lookup"><span data-stu-id="275de-108">OfficeApp</span></span>](officeapp.md)

## <a name="remarks"></a><span data-ttu-id="275de-109">Comentários</span><span class="sxs-lookup"><span data-stu-id="275de-109">Remarks</span></span>

 > [!IMPORTANT]
 > <span data-ttu-id="275de-110">**AllowSnapshot** é `true` por padrão.</span><span class="sxs-lookup"><span data-stu-id="275de-110">**AllowSnapshot** is `true` by default.</span></span> <span data-ttu-id="275de-111">Isso cria uma imagem do suplemento visível para usuários que abrem o documento em uma versão do aplicativo do Office que não oferece suporte a suplementos do Office, ou fornece uma imagem estática do suplemento se o aplicativo não puder se conectar ao servidor que hospeda o suplemento.</span><span class="sxs-lookup"><span data-stu-id="275de-111">This makes an image of the add-in visible for users that open the document in a version of the Office application that doesn't support Office Add-ins, or provides a static image of the add-in if the application can't connect to the server hosting the add-in.</span></span> <span data-ttu-id="275de-112">However, this also means that potentially sensitive information displayed in the add-in can be accessed directly from the document hosting the add-in.</span><span class="sxs-lookup"><span data-stu-id="275de-112">However, this also means that potentially sensitive information displayed in the add-in can be accessed directly from the document hosting the add-in.</span></span>
