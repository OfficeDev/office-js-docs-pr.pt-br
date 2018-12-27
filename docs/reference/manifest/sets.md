---
title: Elemento Sets no arquivo de manifesto
description: ''
ms.date: 10/09/2018
ms.openlocfilehash: b7e78ae05f8409f38c885a1d6a328347d00d0df1
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 12/22/2018
ms.locfileid: "27433652"
---
# <a name="sets-element"></a><span data-ttu-id="e04fe-102">Elemento Sets</span><span class="sxs-lookup"><span data-stu-id="e04fe-102">Sets element</span></span>

<span data-ttu-id="e04fe-103">Especifica o subconjunto mínimo da API do JavaScript para Office que o Suplemento do Office exige para ativar.</span><span class="sxs-lookup"><span data-stu-id="e04fe-103">Specifies the minimum subset of the JavaScript API for Office that your Office Add-in requires in order to activate.</span></span>

<span data-ttu-id="e04fe-104">**Tipo de suplemento:** Conteúdo, Painel de tarefas, Email</span><span class="sxs-lookup"><span data-stu-id="e04fe-104">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="e04fe-105">Sintaxe</span><span class="sxs-lookup"><span data-stu-id="e04fe-105">Syntax</span></span>

```XML
<Sets DefaultMinVersion="n .n ">
   ...
</Sets>
```

## <a name="contained-in"></a><span data-ttu-id="e04fe-106">Contido em</span><span class="sxs-lookup"><span data-stu-id="e04fe-106">Contained in</span></span>

[<span data-ttu-id="e04fe-107">Requisitos</span><span class="sxs-lookup"><span data-stu-id="e04fe-107">Requirements</span></span>](requirements.md)

## <a name="can-contain"></a><span data-ttu-id="e04fe-108">Pode conter</span><span class="sxs-lookup"><span data-stu-id="e04fe-108">Can contain</span></span>

[<span data-ttu-id="e04fe-109">Set</span><span class="sxs-lookup"><span data-stu-id="e04fe-109">Set</span></span>](set.md)

## <a name="attributes"></a><span data-ttu-id="e04fe-110">Atributos</span><span class="sxs-lookup"><span data-stu-id="e04fe-110">Attributes</span></span>

|<span data-ttu-id="e04fe-111">**Atributo**</span><span class="sxs-lookup"><span data-stu-id="e04fe-111">**Attribute**</span></span>|<span data-ttu-id="e04fe-112">**Tipo**</span><span class="sxs-lookup"><span data-stu-id="e04fe-112">**Type**</span></span>|<span data-ttu-id="e04fe-113">**Obrigatório**</span><span class="sxs-lookup"><span data-stu-id="e04fe-113">**Required**</span></span>|<span data-ttu-id="e04fe-114">**Descrição**</span><span class="sxs-lookup"><span data-stu-id="e04fe-114">**Description**</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="e04fe-115">DefaultMinVersion</span><span class="sxs-lookup"><span data-stu-id="e04fe-115">DefaultMinVersion</span></span>|<span data-ttu-id="e04fe-116">cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="e04fe-116">string</span></span>|<span data-ttu-id="e04fe-117">opcional</span><span class="sxs-lookup"><span data-stu-id="e04fe-117">optional</span></span>|<span data-ttu-id="e04fe-p101">Especifica o valor padrão do atributo  **MinVersion** para todos os elementos [Set](set.md) filho. O valor padrão é "1.1".</span><span class="sxs-lookup"><span data-stu-id="e04fe-p101">Specifies the default  **MinVersion** attribute value for all child [Set](set.md) elements. The default value is "1.1".</span></span>|

## <a name="remarks"></a><span data-ttu-id="e04fe-120">Comentários</span><span class="sxs-lookup"><span data-stu-id="e04fe-120">Remarks</span></span>

<span data-ttu-id="e04fe-121">Para saber mais sobre os conjuntos de requisitos, confira [Versões do Office e conjuntos de requisitos](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets).</span><span class="sxs-lookup"><span data-stu-id="e04fe-121">For more information about available requirement sets, see [Office add-in requirement sets](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets).</span></span>

<span data-ttu-id="e04fe-122">Para saber mais sobre o atributo **MinVersion** do elemento **Set** e o atributo **DefaultMinVersion** do elemento **Sets**, confira [Definir o elemento Requirements no manifesto](https://docs.microsoft.com/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements#set-the-requirements-element-in-the-manifest).</span><span class="sxs-lookup"><span data-stu-id="e04fe-122">For more information about the  **MinVersion** attribute of the **Set** element and the **DefaultMinVersion** attribute of the **Sets** element, see [Set the Requirements element in the manifest](https://docs.microsoft.com/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements#set-the-requirements-element-in-the-manifest).</span></span>

