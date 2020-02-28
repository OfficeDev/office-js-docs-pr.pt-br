---
title: Elemento Sets no arquivo de manifesto
description: ''
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: 768f674b4afbd65df88825e871005f182d06f6ce
ms.sourcegitcommit: 5d29801180f6939ec10efb778d2311be67d8b9f1
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 02/27/2020
ms.locfileid: "42325238"
---
# <a name="sets-element"></a><span data-ttu-id="b7dd1-102">Elemento Sets</span><span class="sxs-lookup"><span data-stu-id="b7dd1-102">Sets element</span></span>

<span data-ttu-id="b7dd1-103">Especifica o subconjunto mínimo da API JavaScript do Office que o suplemento do Office exige para ativar.</span><span class="sxs-lookup"><span data-stu-id="b7dd1-103">Specifies the minimum subset of the Office JavaScript API that your Office Add-in requires in order to activate.</span></span>

<span data-ttu-id="b7dd1-104">**Tipo de suplemento:** Conteúdo, Painel de tarefas, Email</span><span class="sxs-lookup"><span data-stu-id="b7dd1-104">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="b7dd1-105">Sintaxe</span><span class="sxs-lookup"><span data-stu-id="b7dd1-105">Syntax</span></span>

```XML
<Sets DefaultMinVersion="n .n ">
   ...
</Sets>
```

## <a name="contained-in"></a><span data-ttu-id="b7dd1-106">Contido em</span><span class="sxs-lookup"><span data-stu-id="b7dd1-106">Contained in</span></span>

[<span data-ttu-id="b7dd1-107">Requisitos</span><span class="sxs-lookup"><span data-stu-id="b7dd1-107">Requirements</span></span>](requirements.md)

## <a name="can-contain"></a><span data-ttu-id="b7dd1-108">Pode conter</span><span class="sxs-lookup"><span data-stu-id="b7dd1-108">Can contain</span></span>

[<span data-ttu-id="b7dd1-109">Set</span><span class="sxs-lookup"><span data-stu-id="b7dd1-109">Set</span></span>](set.md)

## <a name="attributes"></a><span data-ttu-id="b7dd1-110">Atributos</span><span class="sxs-lookup"><span data-stu-id="b7dd1-110">Attributes</span></span>

|<span data-ttu-id="b7dd1-111">**Atributo**</span><span class="sxs-lookup"><span data-stu-id="b7dd1-111">**Attribute**</span></span>|<span data-ttu-id="b7dd1-112">**Tipo**</span><span class="sxs-lookup"><span data-stu-id="b7dd1-112">**Type**</span></span>|<span data-ttu-id="b7dd1-113">**Obrigatório**</span><span class="sxs-lookup"><span data-stu-id="b7dd1-113">**Required**</span></span>|<span data-ttu-id="b7dd1-114">**Descrição**</span><span class="sxs-lookup"><span data-stu-id="b7dd1-114">**Description**</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="b7dd1-115">DefaultMinVersion</span><span class="sxs-lookup"><span data-stu-id="b7dd1-115">DefaultMinVersion</span></span>|<span data-ttu-id="b7dd1-116">cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="b7dd1-116">string</span></span>|<span data-ttu-id="b7dd1-117">opcional</span><span class="sxs-lookup"><span data-stu-id="b7dd1-117">optional</span></span>|<span data-ttu-id="b7dd1-118">Especifica o valor do atributo **MinVersion** padrão para todos os elementos do [conjunto](set.md) filho.</span><span class="sxs-lookup"><span data-stu-id="b7dd1-118">Specifies the default **MinVersion** attribute value for all child [Set](set.md) elements.</span></span> <span data-ttu-id="b7dd1-119">O valor padrão é "1.1".</span><span class="sxs-lookup"><span data-stu-id="b7dd1-119">The default value is "1.1".</span></span>|

## <a name="remarks"></a><span data-ttu-id="b7dd1-120">Comentários</span><span class="sxs-lookup"><span data-stu-id="b7dd1-120">Remarks</span></span>

<span data-ttu-id="b7dd1-121">Para saber mais sobre os conjuntos de requisitos, confira [Versões do Office e conjuntos de requisitos](/office/dev/add-ins/develop/office-versions-and-requirement-sets).</span><span class="sxs-lookup"><span data-stu-id="b7dd1-121">For more information about requirement sets, see [Office versions and requirement sets](/office/dev/add-ins/develop/office-versions-and-requirement-sets).</span></span>

<span data-ttu-id="b7dd1-122">Para obter mais informações sobre o atributo **MinVersion** do elemento **set** e o atributo **DefaultMinVersion** do elemento **sets** , confira [definir o elemento requirements no manifesto](/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements#set-the-requirements-element-in-the-manifest).</span><span class="sxs-lookup"><span data-stu-id="b7dd1-122">For more information about the **MinVersion** attribute of the **Set** element and the **DefaultMinVersion** attribute of the **Sets** element, see [Set the Requirements element in the manifest](/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements#set-the-requirements-element-in-the-manifest).</span></span>

