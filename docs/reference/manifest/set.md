---
title: Elemento Set no arquivo de manifesto
description: O elemento Set especifica um conjunto de requisitos da API JavaScript do Office que o suplemento do Office exige para ativar.
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: e9a70da0dc38c3aee077eb5e7f47cdf8e6dc2d32
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/17/2020
ms.locfileid: "42717912"
---
# <a name="set-element"></a><span data-ttu-id="12fe9-103">Elemento Set</span><span class="sxs-lookup"><span data-stu-id="12fe9-103">Set element</span></span>

<span data-ttu-id="12fe9-104">Especifica um conjunto de requisitos da API JavaScript do Office que o suplemento do Office exige para ativar.</span><span class="sxs-lookup"><span data-stu-id="12fe9-104">Specifies a requirement set from the Office JavaScript API that your Office Add-in requires to activate.</span></span>

<span data-ttu-id="12fe9-105">**Tipo de suplemento:** Conteúdo, Painel de tarefas, Email</span><span class="sxs-lookup"><span data-stu-id="12fe9-105">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="12fe9-106">Sintaxe</span><span class="sxs-lookup"><span data-stu-id="12fe9-106">Syntax</span></span>

```XML
<Set Name="string" MinVersion="n .n">
```

## <a name="contained-in"></a><span data-ttu-id="12fe9-107">Contido em</span><span class="sxs-lookup"><span data-stu-id="12fe9-107">Contained in</span></span>

[<span data-ttu-id="12fe9-108">Sets</span><span class="sxs-lookup"><span data-stu-id="12fe9-108">Sets</span></span>](sets.md)

## <a name="attributes"></a><span data-ttu-id="12fe9-109">Atributos</span><span class="sxs-lookup"><span data-stu-id="12fe9-109">Attributes</span></span>

|<span data-ttu-id="12fe9-110">**Atributo**</span><span class="sxs-lookup"><span data-stu-id="12fe9-110">**Attribute**</span></span>|<span data-ttu-id="12fe9-111">**Tipo**</span><span class="sxs-lookup"><span data-stu-id="12fe9-111">**Type**</span></span>|<span data-ttu-id="12fe9-112">**Obrigatório**</span><span class="sxs-lookup"><span data-stu-id="12fe9-112">**Required**</span></span>|<span data-ttu-id="12fe9-113">**Descrição**</span><span class="sxs-lookup"><span data-stu-id="12fe9-113">**Description**</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="12fe9-114">Nome</span><span class="sxs-lookup"><span data-stu-id="12fe9-114">Name</span></span>|<span data-ttu-id="12fe9-115">string</span><span class="sxs-lookup"><span data-stu-id="12fe9-115">string</span></span>|<span data-ttu-id="12fe9-116">obrigatório</span><span class="sxs-lookup"><span data-stu-id="12fe9-116">required</span></span>|<span data-ttu-id="12fe9-117">O nome de um [conjunto de requisitos](../../develop/office-versions-and-requirement-sets.md).</span><span class="sxs-lookup"><span data-stu-id="12fe9-117">The name of a [requirement set](../../develop/office-versions-and-requirement-sets.md).</span></span>|
|<span data-ttu-id="12fe9-118">MinVersion</span><span class="sxs-lookup"><span data-stu-id="12fe9-118">MinVersion</span></span>|<span data-ttu-id="12fe9-119">cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="12fe9-119">string</span></span>|<span data-ttu-id="12fe9-120">opcional</span><span class="sxs-lookup"><span data-stu-id="12fe9-120">optional</span></span>|<span data-ttu-id="12fe9-121">Especifica a versão mínima do conjunto de APIs exigido pelo seu suplemento.</span><span class="sxs-lookup"><span data-stu-id="12fe9-121">Specifies the minimum version of the API set required by your add-in.</span></span> <span data-ttu-id="12fe9-122">Substitui o valor de **DefaultMinVersion**, se estiver especificado no elemento [sets](sets.md) pai.</span><span class="sxs-lookup"><span data-stu-id="12fe9-122">Overrides the value of **DefaultMinVersion**, if it is specified in the parent [Sets](sets.md) element.</span></span>|

## <a name="remarks"></a><span data-ttu-id="12fe9-123">Comentários</span><span class="sxs-lookup"><span data-stu-id="12fe9-123">Remarks</span></span>

<span data-ttu-id="12fe9-124">Para saber mais sobre os conjuntos de requisitos, confira [Versões do Office e conjuntos de requisitos](../../develop/office-versions-and-requirement-sets.md).</span><span class="sxs-lookup"><span data-stu-id="12fe9-124">For more information about requirement sets, see [Office versions and requirement sets](../../develop/office-versions-and-requirement-sets.md).</span></span>

<span data-ttu-id="12fe9-125">Para obter mais informações sobre o atributo **MinVersion** do elemento **set** e o atributo **DefaultMinVersion** do elemento **sets** , confira [definir o elemento requirements no manifesto](../../develop/specify-office-hosts-and-api-requirements.md#set-the-requirements-element-in-the-manifest).</span><span class="sxs-lookup"><span data-stu-id="12fe9-125">For more information about the **MinVersion** attribute of the **Set** element and the **DefaultMinVersion** attribute of the **Sets** element, see [Set the Requirements element in the manifest](../../develop/specify-office-hosts-and-api-requirements.md#set-the-requirements-element-in-the-manifest).</span></span>

> [!IMPORTANT] 
> <span data-ttu-id="12fe9-126">Para suplementos de email, há apenas um conjunto de requisitos `"Mailbox"` disponível.</span><span class="sxs-lookup"><span data-stu-id="12fe9-126">For mail add-ins, there is only one  `"Mailbox"` requirement set available.</span></span> <span data-ttu-id="12fe9-127">Esse conjunto de requisitos contém o subconjunto completo da API compatível com os suplementos de email do Outlook. Você deve especificar o conjunto de requisitos de `"Mailbox"` no manifesto de seu suplemento de email (não é opcional como no caso de suplementos de conteúdo e do painel de tarefas).</span><span class="sxs-lookup"><span data-stu-id="12fe9-127">This requirement set contains the entire subset of API supported in mail add-ins for Outlook, and you must specify the `"Mailbox"` requirement set in your mail add-in's manifest (it's not optional as is the case for content and task pane add-ins).</span></span> <span data-ttu-id="12fe9-128">Além disso, você não pode declarar suporte para métodos específicos nos suplementos de email.</span><span class="sxs-lookup"><span data-stu-id="12fe9-128">Also, you can't declare support for specific methods in mail add-ins.</span></span>
