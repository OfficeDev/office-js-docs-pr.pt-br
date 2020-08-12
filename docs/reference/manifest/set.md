---
title: Elemento Set no arquivo de manifesto
description: O elemento Set especifica um conjunto de requisitos da API JavaScript do Office que o suplemento do Office exige para ativar.
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: 608830e1ebc0d2e2d4c170b48bba00b3a19e87af
ms.sourcegitcommit: cc6886b47c84ac37a3c957ff85dd0ed526ca5e43
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 08/12/2020
ms.locfileid: "46641414"
---
# <a name="set-element"></a><span data-ttu-id="3a00e-103">Elemento Set</span><span class="sxs-lookup"><span data-stu-id="3a00e-103">Set element</span></span>

<span data-ttu-id="3a00e-104">Especifica um conjunto de requisitos da API JavaScript do Office que o suplemento do Office exige para ativar.</span><span class="sxs-lookup"><span data-stu-id="3a00e-104">Specifies a requirement set from the Office JavaScript API that your Office Add-in requires to activate.</span></span>

<span data-ttu-id="3a00e-105">**Tipo de suplemento:** Conteúdo, Painel de tarefas, Email</span><span class="sxs-lookup"><span data-stu-id="3a00e-105">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="3a00e-106">Sintaxe</span><span class="sxs-lookup"><span data-stu-id="3a00e-106">Syntax</span></span>

```XML
<Set Name="string" MinVersion="n .n">
```

## <a name="contained-in"></a><span data-ttu-id="3a00e-107">Contido em</span><span class="sxs-lookup"><span data-stu-id="3a00e-107">Contained in</span></span>

[<span data-ttu-id="3a00e-108">Sets</span><span class="sxs-lookup"><span data-stu-id="3a00e-108">Sets</span></span>](sets.md)

## <a name="attributes"></a><span data-ttu-id="3a00e-109">Atributos</span><span class="sxs-lookup"><span data-stu-id="3a00e-109">Attributes</span></span>

|<span data-ttu-id="3a00e-110">Atributo</span><span class="sxs-lookup"><span data-stu-id="3a00e-110">Attribute</span></span>|<span data-ttu-id="3a00e-111">Tipo</span><span class="sxs-lookup"><span data-stu-id="3a00e-111">Type</span></span>|<span data-ttu-id="3a00e-112">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="3a00e-112">Required</span></span>|<span data-ttu-id="3a00e-113">Descrição</span><span class="sxs-lookup"><span data-stu-id="3a00e-113">Description</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="3a00e-114">Nome</span><span class="sxs-lookup"><span data-stu-id="3a00e-114">Name</span></span>|<span data-ttu-id="3a00e-115">cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="3a00e-115">string</span></span>|<span data-ttu-id="3a00e-116">obrigatório</span><span class="sxs-lookup"><span data-stu-id="3a00e-116">required</span></span>|<span data-ttu-id="3a00e-117">O nome de um [conjunto de requisitos](../../develop/office-versions-and-requirement-sets.md).</span><span class="sxs-lookup"><span data-stu-id="3a00e-117">The name of a [requirement set](../../develop/office-versions-and-requirement-sets.md).</span></span>|
|<span data-ttu-id="3a00e-118">MinVersion</span><span class="sxs-lookup"><span data-stu-id="3a00e-118">MinVersion</span></span>|<span data-ttu-id="3a00e-119">cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="3a00e-119">string</span></span>|<span data-ttu-id="3a00e-120">opcional</span><span class="sxs-lookup"><span data-stu-id="3a00e-120">optional</span></span>|<span data-ttu-id="3a00e-121">Especifica a versão mínima do conjunto de APIs exigido pelo seu suplemento.</span><span class="sxs-lookup"><span data-stu-id="3a00e-121">Specifies the minimum version of the API set required by your add-in.</span></span> <span data-ttu-id="3a00e-122">Substitui o valor de **DefaultMinVersion**, se estiver especificado no elemento [sets](sets.md) pai.</span><span class="sxs-lookup"><span data-stu-id="3a00e-122">Overrides the value of **DefaultMinVersion**, if it is specified in the parent [Sets](sets.md) element.</span></span>|

## <a name="remarks"></a><span data-ttu-id="3a00e-123">Comentários</span><span class="sxs-lookup"><span data-stu-id="3a00e-123">Remarks</span></span>

<span data-ttu-id="3a00e-124">Para saber mais sobre os conjuntos de requisitos, confira [Versões do Office e conjuntos de requisitos](../../develop/office-versions-and-requirement-sets.md).</span><span class="sxs-lookup"><span data-stu-id="3a00e-124">For more information about requirement sets, see [Office versions and requirement sets](../../develop/office-versions-and-requirement-sets.md).</span></span>

<span data-ttu-id="3a00e-125">Para obter mais informações sobre o atributo **MinVersion** do elemento **set** e o atributo **DefaultMinVersion** do elemento **sets** , confira [definir o elemento requirements no manifesto](../../develop/specify-office-hosts-and-api-requirements.md#set-the-requirements-element-in-the-manifest).</span><span class="sxs-lookup"><span data-stu-id="3a00e-125">For more information about the **MinVersion** attribute of the **Set** element and the **DefaultMinVersion** attribute of the **Sets** element, see [Set the Requirements element in the manifest](../../develop/specify-office-hosts-and-api-requirements.md#set-the-requirements-element-in-the-manifest).</span></span>

> [!IMPORTANT]
> <span data-ttu-id="3a00e-126">Para suplementos de email, há apenas um conjunto de requisitos `"Mailbox"` disponível.</span><span class="sxs-lookup"><span data-stu-id="3a00e-126">For mail add-ins, there is only one  `"Mailbox"` requirement set available.</span></span> <span data-ttu-id="3a00e-127">Esse conjunto de requisitos contém o subconjunto completo da API compatível com os suplementos de email do Outlook. Você deve especificar o conjunto de requisitos de `"Mailbox"` no manifesto de seu suplemento de email (não é opcional como no caso de suplementos de conteúdo e do painel de tarefas).</span><span class="sxs-lookup"><span data-stu-id="3a00e-127">This requirement set contains the entire subset of API supported in mail add-ins for Outlook, and you must specify the `"Mailbox"` requirement set in your mail add-in's manifest (it's not optional as is the case for content and task pane add-ins).</span></span> <span data-ttu-id="3a00e-128">Além disso, você não pode declarar suporte para métodos específicos nos suplementos de email.</span><span class="sxs-lookup"><span data-stu-id="3a00e-128">Also, you can't declare support for specific methods in mail add-ins.</span></span>
