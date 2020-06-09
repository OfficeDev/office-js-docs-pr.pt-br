---
title: Elemento Sets no arquivo de manifesto
description: O elemento sets especifica o conjunto mínimo de API JavaScript do Office que o suplemento do Office exige para ativar.
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: 8c1c97bfc2934ecf3cc20b472b29a03805603729
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 06/08/2020
ms.locfileid: "44608730"
---
# <a name="sets-element"></a><span data-ttu-id="a55f2-103">Elemento Sets</span><span class="sxs-lookup"><span data-stu-id="a55f2-103">Sets element</span></span>

<span data-ttu-id="a55f2-104">Especifica o subconjunto mínimo da API JavaScript do Office que o suplemento do Office exige para ativar.</span><span class="sxs-lookup"><span data-stu-id="a55f2-104">Specifies the minimum subset of the Office JavaScript API that your Office Add-in requires in order to activate.</span></span>

<span data-ttu-id="a55f2-105">**Tipo de suplemento:** Conteúdo, Painel de tarefas, Email</span><span class="sxs-lookup"><span data-stu-id="a55f2-105">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="a55f2-106">Sintaxe</span><span class="sxs-lookup"><span data-stu-id="a55f2-106">Syntax</span></span>

```XML
<Sets DefaultMinVersion="n .n ">
   ...
</Sets>
```

## <a name="contained-in"></a><span data-ttu-id="a55f2-107">Contido em</span><span class="sxs-lookup"><span data-stu-id="a55f2-107">Contained in</span></span>

[<span data-ttu-id="a55f2-108">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a55f2-108">Requirements</span></span>](requirements.md)

## <a name="can-contain"></a><span data-ttu-id="a55f2-109">Pode conter</span><span class="sxs-lookup"><span data-stu-id="a55f2-109">Can contain</span></span>

[<span data-ttu-id="a55f2-110">Set</span><span class="sxs-lookup"><span data-stu-id="a55f2-110">Set</span></span>](set.md)

## <a name="attributes"></a><span data-ttu-id="a55f2-111">Atributos</span><span class="sxs-lookup"><span data-stu-id="a55f2-111">Attributes</span></span>

|<span data-ttu-id="a55f2-112">**Atributo**</span><span class="sxs-lookup"><span data-stu-id="a55f2-112">**Attribute**</span></span>|<span data-ttu-id="a55f2-113">**Tipo**</span><span class="sxs-lookup"><span data-stu-id="a55f2-113">**Type**</span></span>|<span data-ttu-id="a55f2-114">**Obrigatório**</span><span class="sxs-lookup"><span data-stu-id="a55f2-114">**Required**</span></span>|<span data-ttu-id="a55f2-115">**Descrição**</span><span class="sxs-lookup"><span data-stu-id="a55f2-115">**Description**</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="a55f2-116">DefaultMinVersion</span><span class="sxs-lookup"><span data-stu-id="a55f2-116">DefaultMinVersion</span></span>|<span data-ttu-id="a55f2-117">cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="a55f2-117">string</span></span>|<span data-ttu-id="a55f2-118">opcional</span><span class="sxs-lookup"><span data-stu-id="a55f2-118">optional</span></span>|<span data-ttu-id="a55f2-119">Especifica o valor do atributo **MinVersion** padrão para todos os elementos do [conjunto](set.md) filho.</span><span class="sxs-lookup"><span data-stu-id="a55f2-119">Specifies the default **MinVersion** attribute value for all child [Set](set.md) elements.</span></span> <span data-ttu-id="a55f2-120">O valor padrão é "1.1".</span><span class="sxs-lookup"><span data-stu-id="a55f2-120">The default value is "1.1".</span></span>|

## <a name="remarks"></a><span data-ttu-id="a55f2-121">Comentários</span><span class="sxs-lookup"><span data-stu-id="a55f2-121">Remarks</span></span>

<span data-ttu-id="a55f2-122">Para saber mais sobre os conjuntos de requisitos, confira [Versões do Office e conjuntos de requisitos](../../develop/office-versions-and-requirement-sets.md).</span><span class="sxs-lookup"><span data-stu-id="a55f2-122">For more information about requirement sets, see [Office versions and requirement sets](../../develop/office-versions-and-requirement-sets.md).</span></span>

<span data-ttu-id="a55f2-123">Para obter mais informações sobre o atributo **MinVersion** do elemento **set** e o atributo **DefaultMinVersion** do elemento **sets** , confira [definir o elemento requirements no manifesto](../../develop/specify-office-hosts-and-api-requirements.md#set-the-requirements-element-in-the-manifest).</span><span class="sxs-lookup"><span data-stu-id="a55f2-123">For more information about the **MinVersion** attribute of the **Set** element and the **DefaultMinVersion** attribute of the **Sets** element, see [Set the Requirements element in the manifest](../../develop/specify-office-hosts-and-api-requirements.md#set-the-requirements-element-in-the-manifest).</span></span>

