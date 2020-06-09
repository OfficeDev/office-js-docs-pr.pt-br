---
title: Elemento Override no arquivo de manifesto
description: O elemento override permite que você especifique o valor de uma configuração para uma localidade adicional.
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: aa5d023169389670d15e36f8bee4445529d84711
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 06/08/2020
ms.locfileid: "44611502"
---
# <a name="override-element"></a><span data-ttu-id="cd81d-103">Elemento Override</span><span class="sxs-lookup"><span data-stu-id="cd81d-103">Override element</span></span>

<span data-ttu-id="cd81d-104">Fornece uma maneira de especificar o valor de uma configuração para uma localidade adicional.</span><span class="sxs-lookup"><span data-stu-id="cd81d-104">Provides a way to specify the value of a setting for an additional locale.</span></span>

<span data-ttu-id="cd81d-105">**Tipo de suplemento:** Conteúdo, Painel de tarefas, Email</span><span class="sxs-lookup"><span data-stu-id="cd81d-105">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="cd81d-106">Sintaxe</span><span class="sxs-lookup"><span data-stu-id="cd81d-106">Syntax</span></span>

```XML
<Override Locale="string" Value="string" />
```

## <a name="contained-in"></a><span data-ttu-id="cd81d-107">Contido em</span><span class="sxs-lookup"><span data-stu-id="cd81d-107">Contained in</span></span>

|<span data-ttu-id="cd81d-108">**Elemento**</span><span class="sxs-lookup"><span data-stu-id="cd81d-108">**Element**</span></span>|
|:-----|
|[<span data-ttu-id="cd81d-109">CitationText</span><span class="sxs-lookup"><span data-stu-id="cd81d-109">CitationText</span></span>](citationtext.md)|
|[<span data-ttu-id="cd81d-110">Descrição</span><span class="sxs-lookup"><span data-stu-id="cd81d-110">Description</span></span>](description.md)|
|[<span data-ttu-id="cd81d-111">DictionaryName</span><span class="sxs-lookup"><span data-stu-id="cd81d-111">DictionaryName</span></span>](dictionaryname.md)|
|[<span data-ttu-id="cd81d-112">DictionaryHomePage</span><span class="sxs-lookup"><span data-stu-id="cd81d-112">DictionaryHomePage</span></span>](dictionaryhomepage.md)|
|[<span data-ttu-id="cd81d-113">DisplayName</span><span class="sxs-lookup"><span data-stu-id="cd81d-113">DisplayName</span></span>](displayname.md)|
|[<span data-ttu-id="cd81d-114">HighResolutionIconUrl</span><span class="sxs-lookup"><span data-stu-id="cd81d-114">HighResolutionIconUrl</span></span>](highresolutioniconurl.md)|
|[<span data-ttu-id="cd81d-115">IconUrl</span><span class="sxs-lookup"><span data-stu-id="cd81d-115">IconUrl</span></span>](iconurl.md)|
|[<span data-ttu-id="cd81d-116">QueryUri</span><span class="sxs-lookup"><span data-stu-id="cd81d-116">QueryUri</span></span>](queryuri.md)|
|[<span data-ttu-id="cd81d-117">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="cd81d-117">SourceLocation</span></span>](sourcelocation.md)|
|[<span data-ttu-id="cd81d-118">SupportUrl</span><span class="sxs-lookup"><span data-stu-id="cd81d-118">SupportUrl</span></span>](supporturl.md)|

## <a name="attributes"></a><span data-ttu-id="cd81d-119">Atributos</span><span class="sxs-lookup"><span data-stu-id="cd81d-119">Attributes</span></span>

|<span data-ttu-id="cd81d-120">**Atributo**</span><span class="sxs-lookup"><span data-stu-id="cd81d-120">**Attribute**</span></span>|<span data-ttu-id="cd81d-121">**Tipo**</span><span class="sxs-lookup"><span data-stu-id="cd81d-121">**Type**</span></span>|<span data-ttu-id="cd81d-122">**Obrigatório**</span><span class="sxs-lookup"><span data-stu-id="cd81d-122">**Required**</span></span>|<span data-ttu-id="cd81d-123">**Descrição**</span><span class="sxs-lookup"><span data-stu-id="cd81d-123">**Description**</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="cd81d-124">Locale</span><span class="sxs-lookup"><span data-stu-id="cd81d-124">Locale</span></span>|<span data-ttu-id="cd81d-125">cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="cd81d-125">string</span></span>|<span data-ttu-id="cd81d-126">obrigatório</span><span class="sxs-lookup"><span data-stu-id="cd81d-126">required</span></span>|<span data-ttu-id="cd81d-127">Especifica o nome da cultura da localidade para essa substituição no formato de marca do idioma BCP 47, como `"en-US"`.</span><span class="sxs-lookup"><span data-stu-id="cd81d-127">Specifies the culture name of the locale for this override in the BCP 47 language tag format, such as  `"en-US"`.</span></span>|
|<span data-ttu-id="cd81d-128">Valor</span><span class="sxs-lookup"><span data-stu-id="cd81d-128">Value</span></span>|<span data-ttu-id="cd81d-129">cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="cd81d-129">string</span></span>|<span data-ttu-id="cd81d-130">obrigatório</span><span class="sxs-lookup"><span data-stu-id="cd81d-130">required</span></span>|<span data-ttu-id="cd81d-131">Especifica o valor da configuração expressa para a localidade especificada.</span><span class="sxs-lookup"><span data-stu-id="cd81d-131">Specifies value of the setting expressed for the specified locale.</span></span>|

## <a name="see-also"></a><span data-ttu-id="cd81d-132">Confira também</span><span class="sxs-lookup"><span data-stu-id="cd81d-132">See also</span></span>

- [<span data-ttu-id="cd81d-133">Localização para suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="cd81d-133">Localization for Office Add-ins</span></span>](../../develop/localization.md)
