---
title: Elemento AllowSnapshot no arquivo de manifesto
description: ''
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: 02d44167dd1fd46ec6316f3e04393c99f19c9ff0
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 04/24/2019
ms.locfileid: "32450671"
---
# <a name="allowsnapshot-element"></a><span data-ttu-id="98b86-102">Elemento AllowSnapshot</span><span class="sxs-lookup"><span data-stu-id="98b86-102">AllowSnapshot element</span></span>

<span data-ttu-id="98b86-103">Especifica se o instantâneo de uma imagem do suplemento de conteúdo é salvo com o documento host.</span><span class="sxs-lookup"><span data-stu-id="98b86-103">Specifies whether a snapshot image of your content add-in is saved with the host document.</span></span>

<span data-ttu-id="98b86-104">**Tipo de suplemento:** Conteúdo</span><span class="sxs-lookup"><span data-stu-id="98b86-104">**Add-in type:** Content</span></span>

## <a name="syntax"></a><span data-ttu-id="98b86-105">Sintaxe</span><span class="sxs-lookup"><span data-stu-id="98b86-105">Syntax</span></span>

```XML
<AllowSnapshot> [true | false]</AllowSnapshot>
```

## <a name="contained-in"></a><span data-ttu-id="98b86-106">Contido em</span><span class="sxs-lookup"><span data-stu-id="98b86-106">Contained in</span></span>

[<span data-ttu-id="98b86-107">OfficeApp</span><span class="sxs-lookup"><span data-stu-id="98b86-107">OfficeApp</span></span>](officeapp.md)

## <a name="remarks"></a><span data-ttu-id="98b86-108">Comentários</span><span class="sxs-lookup"><span data-stu-id="98b86-108">Remarks</span></span>

 > [!IMPORTANT]
 > <span data-ttu-id="98b86-109">**AllowSnapshot** é `true` por padrão.</span><span class="sxs-lookup"><span data-stu-id="98b86-109">**AllowSnapshot** is `true` by default.</span></span> <span data-ttu-id="98b86-110">Isso cria uma imagem do suplemento visível para os usuários que abrirem o documento em uma versão do aplicativo host que não oferece suporte a Suplementos do Office,ou fornece uma imagem estática do suplemento se o aplicativo host não se conectar ao servidor que hospeda o suplemento.</span><span class="sxs-lookup"><span data-stu-id="98b86-110">This makes an image of the add-in visible for users that open the document in a version of the host application that doesn't support Office Add-ins, or provides a static image of the add-in if the host application can't connect to the server hosting the add-in.</span></span> <span data-ttu-id="98b86-111">However, this also means that potentially sensitive information displayed in the add-in can be accessed directly from the document hosting the add-in.</span><span class="sxs-lookup"><span data-stu-id="98b86-111">However, this also means that potentially sensitive information displayed in the add-in can be accessed directly from the document hosting the add-in.</span></span>

