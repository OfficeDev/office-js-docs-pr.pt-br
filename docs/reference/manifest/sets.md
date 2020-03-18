---
title: Elemento Sets no arquivo de manifesto
description: O elemento sets especifica o conjunto mínimo de API JavaScript do Office que o suplemento do Office exige para ativar.
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: c9e699b4609004c49d954da2367a6c8f82d13670
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/17/2020
ms.locfileid: "42720390"
---
# <a name="sets-element"></a><span data-ttu-id="f3b2c-103">Elemento Sets</span><span class="sxs-lookup"><span data-stu-id="f3b2c-103">Sets element</span></span>

<span data-ttu-id="f3b2c-104">Especifica o subconjunto mínimo da API JavaScript do Office que o suplemento do Office exige para ativar.</span><span class="sxs-lookup"><span data-stu-id="f3b2c-104">Specifies the minimum subset of the Office JavaScript API that your Office Add-in requires in order to activate.</span></span>

<span data-ttu-id="f3b2c-105">**Tipo de suplemento:** Conteúdo, Painel de tarefas, Email</span><span class="sxs-lookup"><span data-stu-id="f3b2c-105">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="f3b2c-106">Sintaxe</span><span class="sxs-lookup"><span data-stu-id="f3b2c-106">Syntax</span></span>

```XML
<Sets DefaultMinVersion="n .n ">
   ...
</Sets>
```

## <a name="contained-in"></a><span data-ttu-id="f3b2c-107">Contido em</span><span class="sxs-lookup"><span data-stu-id="f3b2c-107">Contained in</span></span>

[<span data-ttu-id="f3b2c-108">Requisitos</span><span class="sxs-lookup"><span data-stu-id="f3b2c-108">Requirements</span></span>](requirements.md)

## <a name="can-contain"></a><span data-ttu-id="f3b2c-109">Pode conter</span><span class="sxs-lookup"><span data-stu-id="f3b2c-109">Can contain</span></span>

[<span data-ttu-id="f3b2c-110">Set</span><span class="sxs-lookup"><span data-stu-id="f3b2c-110">Set</span></span>](set.md)

## <a name="attributes"></a><span data-ttu-id="f3b2c-111">Atributos</span><span class="sxs-lookup"><span data-stu-id="f3b2c-111">Attributes</span></span>

|<span data-ttu-id="f3b2c-112">**Atributo**</span><span class="sxs-lookup"><span data-stu-id="f3b2c-112">**Attribute**</span></span>|<span data-ttu-id="f3b2c-113">**Tipo**</span><span class="sxs-lookup"><span data-stu-id="f3b2c-113">**Type**</span></span>|<span data-ttu-id="f3b2c-114">**Obrigatório**</span><span class="sxs-lookup"><span data-stu-id="f3b2c-114">**Required**</span></span>|<span data-ttu-id="f3b2c-115">**Descrição**</span><span class="sxs-lookup"><span data-stu-id="f3b2c-115">**Description**</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="f3b2c-116">DefaultMinVersion</span><span class="sxs-lookup"><span data-stu-id="f3b2c-116">DefaultMinVersion</span></span>|<span data-ttu-id="f3b2c-117">cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="f3b2c-117">string</span></span>|<span data-ttu-id="f3b2c-118">opcional</span><span class="sxs-lookup"><span data-stu-id="f3b2c-118">optional</span></span>|<span data-ttu-id="f3b2c-119">Especifica o valor do atributo **MinVersion** padrão para todos os elementos do [conjunto](set.md) filho.</span><span class="sxs-lookup"><span data-stu-id="f3b2c-119">Specifies the default **MinVersion** attribute value for all child [Set](set.md) elements.</span></span> <span data-ttu-id="f3b2c-120">O valor padrão é "1.1".</span><span class="sxs-lookup"><span data-stu-id="f3b2c-120">The default value is "1.1".</span></span>|

## <a name="remarks"></a><span data-ttu-id="f3b2c-121">Comentários</span><span class="sxs-lookup"><span data-stu-id="f3b2c-121">Remarks</span></span>

<span data-ttu-id="f3b2c-122">Para saber mais sobre os conjuntos de requisitos, confira [Versões do Office e conjuntos de requisitos](../../develop/office-versions-and-requirement-sets.md).</span><span class="sxs-lookup"><span data-stu-id="f3b2c-122">For more information about requirement sets, see [Office versions and requirement sets](../../develop/office-versions-and-requirement-sets.md).</span></span>

<span data-ttu-id="f3b2c-123">Para obter mais informações sobre o atributo **MinVersion** do elemento **set** e o atributo **DefaultMinVersion** do elemento **sets** , confira [definir o elemento requirements no manifesto](../../develop/specify-office-hosts-and-api-requirements.md#set-the-requirements-element-in-the-manifest).</span><span class="sxs-lookup"><span data-stu-id="f3b2c-123">For more information about the **MinVersion** attribute of the **Set** element and the **DefaultMinVersion** attribute of the **Sets** element, see [Set the Requirements element in the manifest](../../develop/specify-office-hosts-and-api-requirements.md#set-the-requirements-element-in-the-manifest).</span></span>

