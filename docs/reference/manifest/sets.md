---
title: Elemento Sets no arquivo de manifesto
description: ''
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: 13777e54ec6bd2d97fa35609ebe194ed85ffa1b8
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 04/24/2019
ms.locfileid: "32450419"
---
# <a name="sets-element"></a><span data-ttu-id="6adfa-102">Elemento Sets</span><span class="sxs-lookup"><span data-stu-id="6adfa-102">Sets element</span></span>

<span data-ttu-id="6adfa-103">Especifica o subconjunto mínimo da API do JavaScript para Office que o Suplemento do Office exige para ativar.</span><span class="sxs-lookup"><span data-stu-id="6adfa-103">Specifies the minimum subset of the JavaScript API for Office that your Office Add-in requires in order to activate.</span></span>

<span data-ttu-id="6adfa-104">**Tipo de suplemento:** Conteúdo, Painel de tarefas, Email</span><span class="sxs-lookup"><span data-stu-id="6adfa-104">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="6adfa-105">Sintaxe</span><span class="sxs-lookup"><span data-stu-id="6adfa-105">Syntax</span></span>

```XML
<Sets DefaultMinVersion="n .n ">
   ...
</Sets>
```

## <a name="contained-in"></a><span data-ttu-id="6adfa-106">Contido em</span><span class="sxs-lookup"><span data-stu-id="6adfa-106">Contained in</span></span>

[<span data-ttu-id="6adfa-107">Requisitos</span><span class="sxs-lookup"><span data-stu-id="6adfa-107">Requirements</span></span>](requirements.md)

## <a name="can-contain"></a><span data-ttu-id="6adfa-108">Pode conter</span><span class="sxs-lookup"><span data-stu-id="6adfa-108">Can contain</span></span>

[<span data-ttu-id="6adfa-109">Set</span><span class="sxs-lookup"><span data-stu-id="6adfa-109">Set</span></span>](set.md)

## <a name="attributes"></a><span data-ttu-id="6adfa-110">Atributos</span><span class="sxs-lookup"><span data-stu-id="6adfa-110">Attributes</span></span>

|<span data-ttu-id="6adfa-111">**Atributo**</span><span class="sxs-lookup"><span data-stu-id="6adfa-111">**Attribute**</span></span>|<span data-ttu-id="6adfa-112">**Tipo**</span><span class="sxs-lookup"><span data-stu-id="6adfa-112">**Type**</span></span>|<span data-ttu-id="6adfa-113">**Obrigatório**</span><span class="sxs-lookup"><span data-stu-id="6adfa-113">**Required**</span></span>|<span data-ttu-id="6adfa-114">**Descrição**</span><span class="sxs-lookup"><span data-stu-id="6adfa-114">**Description**</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="6adfa-115">DefaultMinVersion</span><span class="sxs-lookup"><span data-stu-id="6adfa-115">DefaultMinVersion</span></span>|<span data-ttu-id="6adfa-116">cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="6adfa-116">string</span></span>|<span data-ttu-id="6adfa-117">opcional</span><span class="sxs-lookup"><span data-stu-id="6adfa-117">optional</span></span>|<span data-ttu-id="6adfa-p101">Especifica o valor padrão do atributo  **MinVersion** para todos os elementos [Set](set.md) filho. O valor padrão é "1.1".</span><span class="sxs-lookup"><span data-stu-id="6adfa-p101">Specifies the default  **MinVersion** attribute value for all child [Set](set.md) elements. The default value is "1.1".</span></span>|

## <a name="remarks"></a><span data-ttu-id="6adfa-120">Comentários</span><span class="sxs-lookup"><span data-stu-id="6adfa-120">Remarks</span></span>

<span data-ttu-id="6adfa-121">Para saber mais sobre os conjuntos de requisitos, confira [Versões do Office e conjuntos de requisitos](/office/dev/add-ins/develop/office-versions-and-requirement-sets).</span><span class="sxs-lookup"><span data-stu-id="6adfa-121">For more information about requirement sets, see [Office versions and requirement sets](/office/dev/add-ins/develop/office-versions-and-requirement-sets).</span></span>

<span data-ttu-id="6adfa-122">Para saber mais sobre o atributo **MinVersion** do elemento **Set** e o atributo **DefaultMinVersion** do elemento **Sets**, confira [Definir o elemento Requirements no manifesto](/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements#set-the-requirements-element-in-the-manifest).</span><span class="sxs-lookup"><span data-stu-id="6adfa-122">For more information about the  **MinVersion** attribute of the **Set** element and the **DefaultMinVersion** attribute of the **Sets** element, see [Set the Requirements element in the manifest](/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements#set-the-requirements-element-in-the-manifest).</span></span>

