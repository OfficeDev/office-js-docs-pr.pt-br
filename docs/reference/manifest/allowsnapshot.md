---
title: Elemento AllowSnapshot no arquivo de manifesto
description: Especifica se o instantâneo de uma imagem do suplemento de conteúdo é salvo com o documento host.
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: 8bb143d13a17b3e184af64f1bf18f2a32a55b60c
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/17/2020
ms.locfileid: "42720957"
---
# <a name="allowsnapshot-element"></a><span data-ttu-id="09185-103">Elemento AllowSnapshot</span><span class="sxs-lookup"><span data-stu-id="09185-103">AllowSnapshot element</span></span>

<span data-ttu-id="09185-104">Especifica se o instantâneo de uma imagem do suplemento de conteúdo é salvo com o documento host.</span><span class="sxs-lookup"><span data-stu-id="09185-104">Specifies whether a snapshot image of your content add-in is saved with the host document.</span></span>

<span data-ttu-id="09185-105">**Tipo de suplemento:** Conteúdo</span><span class="sxs-lookup"><span data-stu-id="09185-105">**Add-in type:** Content</span></span>

## <a name="syntax"></a><span data-ttu-id="09185-106">Sintaxe</span><span class="sxs-lookup"><span data-stu-id="09185-106">Syntax</span></span>

```XML
<AllowSnapshot> [true | false]</AllowSnapshot>
```

## <a name="contained-in"></a><span data-ttu-id="09185-107">Contido em</span><span class="sxs-lookup"><span data-stu-id="09185-107">Contained in</span></span>

[<span data-ttu-id="09185-108">OfficeApp</span><span class="sxs-lookup"><span data-stu-id="09185-108">OfficeApp</span></span>](officeapp.md)

## <a name="remarks"></a><span data-ttu-id="09185-109">Comentários</span><span class="sxs-lookup"><span data-stu-id="09185-109">Remarks</span></span>

 > [!IMPORTANT]
 > <span data-ttu-id="09185-110">**AllowSnapshot** é `true` por padrão.</span><span class="sxs-lookup"><span data-stu-id="09185-110">**AllowSnapshot** is `true` by default.</span></span> <span data-ttu-id="09185-111">Isso cria uma imagem do suplemento visível para os usuários que abrirem o documento em uma versão do aplicativo host que não oferece suporte a Suplementos do Office,ou fornece uma imagem estática do suplemento se o aplicativo host não se conectar ao servidor que hospeda o suplemento.</span><span class="sxs-lookup"><span data-stu-id="09185-111">This makes an image of the add-in visible for users that open the document in a version of the host application that doesn't support Office Add-ins, or provides a static image of the add-in if the host application can't connect to the server hosting the add-in.</span></span> <span data-ttu-id="09185-112">However, this also means that potentially sensitive information displayed in the add-in can be accessed directly from the document hosting the add-in.</span><span class="sxs-lookup"><span data-stu-id="09185-112">However, this also means that potentially sensitive information displayed in the add-in can be accessed directly from the document hosting the add-in.</span></span>

