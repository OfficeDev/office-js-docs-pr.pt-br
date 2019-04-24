---
title: Elemento Set no arquivo de manifesto
description: ''
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: 0f408d698d297eaa6287ff268bdb7fc737a5a24d
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 04/24/2019
ms.locfileid: "32452029"
---
# <a name="set-element"></a><span data-ttu-id="1d292-102">Elemento Set</span><span class="sxs-lookup"><span data-stu-id="1d292-102">Set element</span></span>

<span data-ttu-id="1d292-103">Especifica um conjunto de requisitos a partir da API do JavaScript para Office que o seu Suplemento do Office exige para ativar.</span><span class="sxs-lookup"><span data-stu-id="1d292-103">Specifies a requirement set from the JavaScript API for Office that your Office Add-in requires to activate.</span></span>

<span data-ttu-id="1d292-104">**Tipo de suplemento:** Conteúdo, Painel de tarefas, Email</span><span class="sxs-lookup"><span data-stu-id="1d292-104">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="1d292-105">Sintaxe</span><span class="sxs-lookup"><span data-stu-id="1d292-105">Syntax</span></span>

```XML
<Set Name="string" MinVersion="n .n">
```

## <a name="contained-in"></a><span data-ttu-id="1d292-106">Contido em</span><span class="sxs-lookup"><span data-stu-id="1d292-106">Contained in</span></span>

[<span data-ttu-id="1d292-107">Sets</span><span class="sxs-lookup"><span data-stu-id="1d292-107">Sets</span></span>](sets.md)

## <a name="attributes"></a><span data-ttu-id="1d292-108">Atributos</span><span class="sxs-lookup"><span data-stu-id="1d292-108">Attributes</span></span>

|<span data-ttu-id="1d292-109">**Atributo**</span><span class="sxs-lookup"><span data-stu-id="1d292-109">**Attribute**</span></span>|<span data-ttu-id="1d292-110">**Tipo**</span><span class="sxs-lookup"><span data-stu-id="1d292-110">**Type**</span></span>|<span data-ttu-id="1d292-111">**Obrigatório**</span><span class="sxs-lookup"><span data-stu-id="1d292-111">**Required**</span></span>|<span data-ttu-id="1d292-112">**Descrição**</span><span class="sxs-lookup"><span data-stu-id="1d292-112">**Description**</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="1d292-113">Nome</span><span class="sxs-lookup"><span data-stu-id="1d292-113">Name</span></span>|<span data-ttu-id="1d292-114">string</span><span class="sxs-lookup"><span data-stu-id="1d292-114">string</span></span>|<span data-ttu-id="1d292-115">obrigatório</span><span class="sxs-lookup"><span data-stu-id="1d292-115">required</span></span>|<span data-ttu-id="1d292-116">O nome de um [conjunto de requisitos](/office/dev/add-ins/develop/office-versions-and-requirement-sets).</span><span class="sxs-lookup"><span data-stu-id="1d292-116">The name of a [requirement set](/office/dev/add-ins/develop/office-versions-and-requirement-sets).</span></span>|
|<span data-ttu-id="1d292-117">MinVersion</span><span class="sxs-lookup"><span data-stu-id="1d292-117">MinVersion</span></span>|<span data-ttu-id="1d292-118">cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="1d292-118">string</span></span>|<span data-ttu-id="1d292-119">opcional</span><span class="sxs-lookup"><span data-stu-id="1d292-119">optional</span></span>|<span data-ttu-id="1d292-p101">Especifica a versão mínima do conjunto de APIs exigido pelo seu suplemento. Substitui o valor de **DefaultMinVersion**, se ele estiver especificado no elemento [Sets](sets.md) pai.</span><span class="sxs-lookup"><span data-stu-id="1d292-p101">Specifies the minimum version of the API set required by your add-in. Overrides the value of  **DefaultMinVersion**, if it is specified in the parent [Sets](sets.md) element.</span></span>|

## <a name="remarks"></a><span data-ttu-id="1d292-122">Comentários</span><span class="sxs-lookup"><span data-stu-id="1d292-122">Remarks</span></span>

<span data-ttu-id="1d292-123">Para saber mais sobre os conjuntos de requisitos, confira [Versões do Office e conjuntos de requisitos](/office/dev/add-ins/develop/office-versions-and-requirement-sets).</span><span class="sxs-lookup"><span data-stu-id="1d292-123">For more information about requirement sets, see [Office versions and requirement sets](/office/dev/add-ins/develop/office-versions-and-requirement-sets).</span></span>

<span data-ttu-id="1d292-124">Para saber mais sobre o atributo **MinVersion** do elemento **Set** e o atributo **DefaultMinVersion** do elemento **Sets**, confira [Definir o elemento Requirements no manifesto](/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements#set-the-requirements-element-in-the-manifest).</span><span class="sxs-lookup"><span data-stu-id="1d292-124">For more information about the  **MinVersion** attribute of the **Set** element and the **DefaultMinVersion** attribute of the **Sets** element, see [Set the Requirements element in the manifest](/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements#set-the-requirements-element-in-the-manifest).</span></span>

> [!IMPORTANT] 
> <span data-ttu-id="1d292-125">Para suplementos de email, há apenas um conjunto de requisitos `"Mailbox"` disponível.</span><span class="sxs-lookup"><span data-stu-id="1d292-125">For mail add-ins, there is only one  `"Mailbox"` requirement set available.</span></span> <span data-ttu-id="1d292-126">Esse conjunto de requisitos contém o subconjunto completo da API compatível com os suplementos de email do Outlook. Você deve especificar o conjunto de requisitos de `"Mailbox"` no manifesto de seu suplemento de email (não é opcional como no caso de suplementos de conteúdo e do painel de tarefas).</span><span class="sxs-lookup"><span data-stu-id="1d292-126">This requirement set contains the entire subset of API supported in mail add-ins for Outlook, and you must specify the `"Mailbox"` requirement set in your mail add-in's manifest (it's not optional as is the case for content and task pane add-ins).</span></span> <span data-ttu-id="1d292-127">Além disso, você não pode declarar suporte para métodos específicos nos suplementos de email.</span><span class="sxs-lookup"><span data-stu-id="1d292-127">Also, you can't declare support for specific methods in mail add-ins.</span></span>
