---
title: Elemento AllowSnapshot no arquivo de manifesto
description: ''
ms.date: 10/09/2018
ms.openlocfilehash: f1aced0ce37b01c277ea5a8621f6c7764d2f761b
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 12/22/2018
ms.locfileid: "27432344"
---
# <a name="allowsnapshot-element"></a><span data-ttu-id="3ae45-102">Elemento AllowSnapshot</span><span class="sxs-lookup"><span data-stu-id="3ae45-102">AllowSnapshot element</span></span>

<span data-ttu-id="3ae45-103">Especifica se o instantâneo de uma imagem do suplemento de conteúdo é salvo com o documento host.</span><span class="sxs-lookup"><span data-stu-id="3ae45-103">Specifies whether a snapshot image of your content add-in is saved with the host document.</span></span>

<span data-ttu-id="3ae45-104">**Tipo de suplemento:** Conteúdo</span><span class="sxs-lookup"><span data-stu-id="3ae45-104">**Add-in type:** Content</span></span>

## <a name="syntax"></a><span data-ttu-id="3ae45-105">Sintaxe</span><span class="sxs-lookup"><span data-stu-id="3ae45-105">Syntax</span></span>

```XML
<AllowSnapshot> [true | false]</AllowSnapshot>
```

## <a name="contained-in"></a><span data-ttu-id="3ae45-106">Contido em</span><span class="sxs-lookup"><span data-stu-id="3ae45-106">Contained in</span></span>

[<span data-ttu-id="3ae45-107">OfficeApp</span><span class="sxs-lookup"><span data-stu-id="3ae45-107">OfficeApp</span></span>](officeapp.md)

## <a name="remarks"></a><span data-ttu-id="3ae45-108">Comentários</span><span class="sxs-lookup"><span data-stu-id="3ae45-108">Remarks</span></span>

 > [!IMPORTANT]
 > <span data-ttu-id="3ae45-109">**AllowSnapshot** é `true` por padrão.</span><span class="sxs-lookup"><span data-stu-id="3ae45-109">Security Note:**AllowSnapshot** is true`true` by default.</span></span> <span data-ttu-id="3ae45-110">Isso cria uma imagem do suplemento visível para os usuários que abrirem o documento em uma versão do aplicativo host que não oferece suporte a Suplementos do Office,ou fornece uma imagem estática do suplemento se o aplicativo host não se conectar ao servidor que hospeda o suplemento.</span><span class="sxs-lookup"><span data-stu-id="3ae45-110">This makes an image of the add-in visible for users that open the document in a version of the host application that doesn't support Office Add-ins, or provides a static image of the add-in if the host application can't connect to the server hosting the add-in.</span></span> <span data-ttu-id="3ae45-111">No entanto, isso também significa que informações potencialmente confidenciais exibidas no suplemento podem ser acessadas diretamente no documento que hospeda o suplemento.</span><span class="sxs-lookup"><span data-stu-id="3ae45-111">However, this also means that potentially sensitive information displayed in the add-in can be accessed directly from the document hosting the add-in.</span></span>

