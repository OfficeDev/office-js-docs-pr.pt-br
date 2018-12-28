---
title: Elemento Set no arquivo de manifesto
description: ''
ms.date: 10/09/2018
ms.openlocfilehash: 0f137f7b08d6f1d0b0d972173c8085713b0f979d
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 12/22/2018
ms.locfileid: "27432764"
---
# <a name="set-element"></a><span data-ttu-id="56adf-102">Elemento Set</span><span class="sxs-lookup"><span data-stu-id="56adf-102">Set element</span></span>

<span data-ttu-id="56adf-103">Especifica um conjunto de requisitos a partir da API do JavaScript para Office que o seu Suplemento do Office exige para ativar.</span><span class="sxs-lookup"><span data-stu-id="56adf-103">Specifies a requirement set from the JavaScript API for Office that your Office Add-in requires to activate.</span></span>

<span data-ttu-id="56adf-104">**Tipo de suplemento:** Conteúdo, Painel de tarefas, Email</span><span class="sxs-lookup"><span data-stu-id="56adf-104">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="56adf-105">Sintaxe</span><span class="sxs-lookup"><span data-stu-id="56adf-105">Syntax</span></span>

```XML
<Set Name="string" MinVersion="n .n">
```

## <a name="contained-in"></a><span data-ttu-id="56adf-106">Contido em</span><span class="sxs-lookup"><span data-stu-id="56adf-106">Contained in</span></span>

[<span data-ttu-id="56adf-107">Sets</span><span class="sxs-lookup"><span data-stu-id="56adf-107">Sets</span></span>](sets.md)

## <a name="attributes"></a><span data-ttu-id="56adf-108">Atributos</span><span class="sxs-lookup"><span data-stu-id="56adf-108">Attributes</span></span>

|<span data-ttu-id="56adf-109">**Atributo**</span><span class="sxs-lookup"><span data-stu-id="56adf-109">**Attribute**</span></span>|<span data-ttu-id="56adf-110">**Tipo**</span><span class="sxs-lookup"><span data-stu-id="56adf-110">**Type**</span></span>|<span data-ttu-id="56adf-111">**Obrigatório**</span><span class="sxs-lookup"><span data-stu-id="56adf-111">**Required**</span></span>|<span data-ttu-id="56adf-112">**Descrição**</span><span class="sxs-lookup"><span data-stu-id="56adf-112">**Description**</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="56adf-113">Nome</span><span class="sxs-lookup"><span data-stu-id="56adf-113">Name</span></span>|<span data-ttu-id="56adf-114">cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="56adf-114">string</span></span>|<span data-ttu-id="56adf-115">obrigatório</span><span class="sxs-lookup"><span data-stu-id="56adf-115">required</span></span>|<span data-ttu-id="56adf-116">O nome de um [conjunto de requisitos](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets).</span><span class="sxs-lookup"><span data-stu-id="56adf-116">The name of a [requirement set](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets).</span></span>|
|<span data-ttu-id="56adf-117">MinVersion</span><span class="sxs-lookup"><span data-stu-id="56adf-117">MinVersion</span></span>|<span data-ttu-id="56adf-118">cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="56adf-118">string</span></span>|<span data-ttu-id="56adf-119">opcional</span><span class="sxs-lookup"><span data-stu-id="56adf-119">optional</span></span>|<span data-ttu-id="56adf-p101">Especifica a versão mínima do conjunto de APIs exigido pelo seu suplemento. Substitui o valor de **DefaultMinVersion**, se ele estiver especificado no elemento [Sets](sets.md) pai.</span><span class="sxs-lookup"><span data-stu-id="56adf-p101">Specifies the minimum version of the API set required by your add-in. Overrides the value of  **DefaultMinVersion**, if it is specified in the parent [Sets](sets.md) element.</span></span>|

## <a name="remarks"></a><span data-ttu-id="56adf-122">Comentários</span><span class="sxs-lookup"><span data-stu-id="56adf-122">Remarks</span></span>

<span data-ttu-id="56adf-123">Para saber mais sobre os conjuntos de requisitos, confira [Versões do Office e conjuntos de requisitos](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets).</span><span class="sxs-lookup"><span data-stu-id="56adf-123">For more information about available requirement sets, see [Office add-in requirement sets](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets).</span></span>

<span data-ttu-id="56adf-124">Para saber mais sobre o atributo **MinVersion** do elemento **Set** e o atributo **DefaultMinVersion** do elemento **Sets**, confira [Definir o elemento Requirements no manifesto](https://docs.microsoft.com/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements#set-the-requirements-element-in-the-manifest).</span><span class="sxs-lookup"><span data-stu-id="56adf-124">For more information about the  **MinVersion** attribute of the **Set** element and the **DefaultMinVersion** attribute of the **Sets** element, see [Set the Requirements element in the manifest](https://docs.microsoft.com/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements#set-the-requirements-element-in-the-manifest).</span></span>

> [!IMPORTANT] 
> <span data-ttu-id="56adf-125">Para suplementos de email, há apenas um conjunto de requisitos `"Mailbox"` disponível.</span><span class="sxs-lookup"><span data-stu-id="56adf-125">For mail add-ins, there is only one  `"Mailbox"` requirement set available.</span></span> <span data-ttu-id="56adf-126">Esse conjunto de requisitos contém o subconjunto completo da API compatível com os suplementos de email do Outlook. Você deve especificar o conjunto de requisitos de `"Mailbox"` no manifesto de seu suplemento de email (não é opcional como no caso de suplementos de conteúdo e do painel de tarefas).</span><span class="sxs-lookup"><span data-stu-id="56adf-126">Important  For mail add-ins, there is only one   requirement set available. This requirement set contains the entire subset of API supported in mail add-ins for Outlook, and you must specify the `"Mailbox"` requirement set in your mail add-in's manifest (it's not optional as is the case for content and task pane add-ins). Also, you can't declare support for specific methods in mail add-ins.</span></span> <span data-ttu-id="56adf-127">Além disso, não é possível declarar suporte para métodos específicos em suplementos de email.</span><span class="sxs-lookup"><span data-stu-id="56adf-127">Also, you can't declare support for specific methods in mail add-ins.</span></span>
