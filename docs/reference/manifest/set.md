---
title: Elemento Set no arquivo de manifesto
description: ''
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: 47f675f999a225e499171cb03c27797bb3dcc5f6
ms.sourcegitcommit: 4079903c3cc45b7d8c041509a44e9fc38da399b1
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/11/2020
ms.locfileid: "42596500"
---
# <a name="set-element"></a><span data-ttu-id="74070-102">Elemento Set</span><span class="sxs-lookup"><span data-stu-id="74070-102">Set element</span></span>

<span data-ttu-id="74070-103">Especifica um conjunto de requisitos da API JavaScript do Office que o suplemento do Office exige para ativar.</span><span class="sxs-lookup"><span data-stu-id="74070-103">Specifies a requirement set from the Office JavaScript API that your Office Add-in requires to activate.</span></span>

<span data-ttu-id="74070-104">**Tipo de suplemento:** Conteúdo, Painel de tarefas, Email</span><span class="sxs-lookup"><span data-stu-id="74070-104">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="74070-105">Sintaxe</span><span class="sxs-lookup"><span data-stu-id="74070-105">Syntax</span></span>

```XML
<Set Name="string" MinVersion="n .n">
```

## <a name="contained-in"></a><span data-ttu-id="74070-106">Contido em</span><span class="sxs-lookup"><span data-stu-id="74070-106">Contained in</span></span>

[<span data-ttu-id="74070-107">Sets</span><span class="sxs-lookup"><span data-stu-id="74070-107">Sets</span></span>](sets.md)

## <a name="attributes"></a><span data-ttu-id="74070-108">Atributos</span><span class="sxs-lookup"><span data-stu-id="74070-108">Attributes</span></span>

|<span data-ttu-id="74070-109">**Atributo**</span><span class="sxs-lookup"><span data-stu-id="74070-109">**Attribute**</span></span>|<span data-ttu-id="74070-110">**Tipo**</span><span class="sxs-lookup"><span data-stu-id="74070-110">**Type**</span></span>|<span data-ttu-id="74070-111">**Obrigatório**</span><span class="sxs-lookup"><span data-stu-id="74070-111">**Required**</span></span>|<span data-ttu-id="74070-112">**Descrição**</span><span class="sxs-lookup"><span data-stu-id="74070-112">**Description**</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="74070-113">Nome</span><span class="sxs-lookup"><span data-stu-id="74070-113">Name</span></span>|<span data-ttu-id="74070-114">cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="74070-114">string</span></span>|<span data-ttu-id="74070-115">obrigatório</span><span class="sxs-lookup"><span data-stu-id="74070-115">required</span></span>|<span data-ttu-id="74070-116">O nome de um [conjunto de requisitos](../../develop/office-versions-and-requirement-sets.md).</span><span class="sxs-lookup"><span data-stu-id="74070-116">The name of a [requirement set](../../develop/office-versions-and-requirement-sets.md).</span></span>|
|<span data-ttu-id="74070-117">MinVersion</span><span class="sxs-lookup"><span data-stu-id="74070-117">MinVersion</span></span>|<span data-ttu-id="74070-118">cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="74070-118">string</span></span>|<span data-ttu-id="74070-119">opcional</span><span class="sxs-lookup"><span data-stu-id="74070-119">optional</span></span>|<span data-ttu-id="74070-120">Especifica a versão mínima do conjunto de APIs exigido pelo seu suplemento.</span><span class="sxs-lookup"><span data-stu-id="74070-120">Specifies the minimum version of the API set required by your add-in.</span></span> <span data-ttu-id="74070-121">Substitui o valor de **DefaultMinVersion**, se estiver especificado no elemento [sets](sets.md) pai.</span><span class="sxs-lookup"><span data-stu-id="74070-121">Overrides the value of **DefaultMinVersion**, if it is specified in the parent [Sets](sets.md) element.</span></span>|

## <a name="remarks"></a><span data-ttu-id="74070-122">Comentários</span><span class="sxs-lookup"><span data-stu-id="74070-122">Remarks</span></span>

<span data-ttu-id="74070-123">Para saber mais sobre os conjuntos de requisitos, confira [Versões do Office e conjuntos de requisitos](../../develop/office-versions-and-requirement-sets.md).</span><span class="sxs-lookup"><span data-stu-id="74070-123">For more information about requirement sets, see [Office versions and requirement sets](../../develop/office-versions-and-requirement-sets.md).</span></span>

<span data-ttu-id="74070-124">Para obter mais informações sobre o atributo **MinVersion** do elemento **set** e o atributo **DefaultMinVersion** do elemento **sets** , confira [definir o elemento requirements no manifesto](../../develop/specify-office-hosts-and-api-requirements.md#set-the-requirements-element-in-the-manifest).</span><span class="sxs-lookup"><span data-stu-id="74070-124">For more information about the **MinVersion** attribute of the **Set** element and the **DefaultMinVersion** attribute of the **Sets** element, see [Set the Requirements element in the manifest](../../develop/specify-office-hosts-and-api-requirements.md#set-the-requirements-element-in-the-manifest).</span></span>

> [!IMPORTANT] 
> <span data-ttu-id="74070-125">Para suplementos de email, há apenas um conjunto de requisitos `"Mailbox"` disponível.</span><span class="sxs-lookup"><span data-stu-id="74070-125">For mail add-ins, there is only one  `"Mailbox"` requirement set available.</span></span> <span data-ttu-id="74070-126">Esse conjunto de requisitos contém o subconjunto completo da API compatível com os suplementos de email do Outlook. Você deve especificar o conjunto de requisitos de `"Mailbox"` no manifesto de seu suplemento de email (não é opcional como no caso de suplementos de conteúdo e do painel de tarefas).</span><span class="sxs-lookup"><span data-stu-id="74070-126">This requirement set contains the entire subset of API supported in mail add-ins for Outlook, and you must specify the `"Mailbox"` requirement set in your mail add-in's manifest (it's not optional as is the case for content and task pane add-ins).</span></span> <span data-ttu-id="74070-127">Além disso, você não pode declarar suporte para métodos específicos nos suplementos de email.</span><span class="sxs-lookup"><span data-stu-id="74070-127">Also, you can't declare support for specific methods in mail add-ins.</span></span>
