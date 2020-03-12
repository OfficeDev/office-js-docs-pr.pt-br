---
title: Elemento Override no arquivo de manifesto
description: ''
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: a1e11257e28d015d6fca9c9a1868e75989616e16
ms.sourcegitcommit: 4079903c3cc45b7d8c041509a44e9fc38da399b1
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/11/2020
ms.locfileid: "42596876"
---
# <a name="override-element"></a><span data-ttu-id="ec5ac-102">Elemento Override</span><span class="sxs-lookup"><span data-stu-id="ec5ac-102">Override element</span></span>

<span data-ttu-id="ec5ac-103">Fornece uma maneira de especificar o valor de uma configuração para uma localidade adicional.</span><span class="sxs-lookup"><span data-stu-id="ec5ac-103">Provides a way to specify the value of a setting for an additional locale.</span></span>

<span data-ttu-id="ec5ac-104">**Tipo de suplemento:** Conteúdo, Painel de tarefas, Email</span><span class="sxs-lookup"><span data-stu-id="ec5ac-104">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="ec5ac-105">Sintaxe</span><span class="sxs-lookup"><span data-stu-id="ec5ac-105">Syntax</span></span>

```XML
<Override Locale="string" Value="string" />
```

## <a name="contained-in"></a><span data-ttu-id="ec5ac-106">Contido em</span><span class="sxs-lookup"><span data-stu-id="ec5ac-106">Contained in</span></span>

|<span data-ttu-id="ec5ac-107">**Elemento**</span><span class="sxs-lookup"><span data-stu-id="ec5ac-107">**Element**</span></span>|
|:-----|
|[<span data-ttu-id="ec5ac-108">CitationText</span><span class="sxs-lookup"><span data-stu-id="ec5ac-108">CitationText</span></span>](citationtext.md)|
|[<span data-ttu-id="ec5ac-109">Descrição</span><span class="sxs-lookup"><span data-stu-id="ec5ac-109">Description</span></span>](description.md)|
|[<span data-ttu-id="ec5ac-110">DictionaryName</span><span class="sxs-lookup"><span data-stu-id="ec5ac-110">DictionaryName</span></span>](dictionaryname.md)|
|[<span data-ttu-id="ec5ac-111">DictionaryHomePage</span><span class="sxs-lookup"><span data-stu-id="ec5ac-111">DictionaryHomePage</span></span>](dictionaryhomepage.md)|
|[<span data-ttu-id="ec5ac-112">DisplayName</span><span class="sxs-lookup"><span data-stu-id="ec5ac-112">DisplayName</span></span>](displayname.md)|
|[<span data-ttu-id="ec5ac-113">HighResolutionIconUrl</span><span class="sxs-lookup"><span data-stu-id="ec5ac-113">HighResolutionIconUrl</span></span>](highresolutioniconurl.md)|
|[<span data-ttu-id="ec5ac-114">IconUrl</span><span class="sxs-lookup"><span data-stu-id="ec5ac-114">IconUrl</span></span>](iconurl.md)|
|[<span data-ttu-id="ec5ac-115">QueryUri</span><span class="sxs-lookup"><span data-stu-id="ec5ac-115">QueryUri</span></span>](queryuri.md)|
|[<span data-ttu-id="ec5ac-116">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="ec5ac-116">SourceLocation</span></span>](sourcelocation.md)|
|[<span data-ttu-id="ec5ac-117">SupportUrl</span><span class="sxs-lookup"><span data-stu-id="ec5ac-117">SupportUrl</span></span>](supporturl.md)|

## <a name="attributes"></a><span data-ttu-id="ec5ac-118">Atributos</span><span class="sxs-lookup"><span data-stu-id="ec5ac-118">Attributes</span></span>

|<span data-ttu-id="ec5ac-119">**Atributo**</span><span class="sxs-lookup"><span data-stu-id="ec5ac-119">**Attribute**</span></span>|<span data-ttu-id="ec5ac-120">**Tipo**</span><span class="sxs-lookup"><span data-stu-id="ec5ac-120">**Type**</span></span>|<span data-ttu-id="ec5ac-121">**Obrigatório**</span><span class="sxs-lookup"><span data-stu-id="ec5ac-121">**Required**</span></span>|<span data-ttu-id="ec5ac-122">**Descrição**</span><span class="sxs-lookup"><span data-stu-id="ec5ac-122">**Description**</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="ec5ac-123">Locale</span><span class="sxs-lookup"><span data-stu-id="ec5ac-123">Locale</span></span>|<span data-ttu-id="ec5ac-124">cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="ec5ac-124">string</span></span>|<span data-ttu-id="ec5ac-125">obrigatório</span><span class="sxs-lookup"><span data-stu-id="ec5ac-125">required</span></span>|<span data-ttu-id="ec5ac-126">Especifica o nome da cultura da localidade para essa substituição no formato de marca do idioma BCP 47, como `"en-US"`.</span><span class="sxs-lookup"><span data-stu-id="ec5ac-126">Specifies the culture name of the locale for this override in the BCP 47 language tag format, such as  `"en-US"`.</span></span>|
|<span data-ttu-id="ec5ac-127">Valor</span><span class="sxs-lookup"><span data-stu-id="ec5ac-127">Value</span></span>|<span data-ttu-id="ec5ac-128">cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="ec5ac-128">string</span></span>|<span data-ttu-id="ec5ac-129">obrigatório</span><span class="sxs-lookup"><span data-stu-id="ec5ac-129">required</span></span>|<span data-ttu-id="ec5ac-130">Especifica o valor da configuração expressa para a localidade especificada.</span><span class="sxs-lookup"><span data-stu-id="ec5ac-130">Specifies value of the setting expressed for the specified locale.</span></span>|

## <a name="see-also"></a><span data-ttu-id="ec5ac-131">Confira também</span><span class="sxs-lookup"><span data-stu-id="ec5ac-131">See also</span></span>

- [<span data-ttu-id="ec5ac-132">Localização para suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="ec5ac-132">Localization for Office Add-ins</span></span>](../../develop/localization.md)
