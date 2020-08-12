---
title: Elemento Override no arquivo de manifesto
description: O elemento override permite que você especifique o valor de uma configuração para uma localidade adicional.
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: 139a4089a36d8a8adfa71d4a0947b02f5b163b52
ms.sourcegitcommit: cc6886b47c84ac37a3c957ff85dd0ed526ca5e43
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 08/12/2020
ms.locfileid: "46641449"
---
# <a name="override-element"></a><span data-ttu-id="ca14e-103">Elemento Override</span><span class="sxs-lookup"><span data-stu-id="ca14e-103">Override element</span></span>

<span data-ttu-id="ca14e-104">Fornece uma maneira de especificar o valor de uma configuração para uma localidade adicional.</span><span class="sxs-lookup"><span data-stu-id="ca14e-104">Provides a way to specify the value of a setting for an additional locale.</span></span>

<span data-ttu-id="ca14e-105">**Tipo de suplemento:** Conteúdo, Painel de tarefas, Email</span><span class="sxs-lookup"><span data-stu-id="ca14e-105">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="ca14e-106">Sintaxe</span><span class="sxs-lookup"><span data-stu-id="ca14e-106">Syntax</span></span>

```XML
<Override Locale="string" Value="string" />
```

## <a name="contained-in"></a><span data-ttu-id="ca14e-107">Contido em</span><span class="sxs-lookup"><span data-stu-id="ca14e-107">Contained in</span></span>

|<span data-ttu-id="ca14e-108">Elemento</span><span class="sxs-lookup"><span data-stu-id="ca14e-108">Element</span></span>|
|:-----|
|[<span data-ttu-id="ca14e-109">CitationText</span><span class="sxs-lookup"><span data-stu-id="ca14e-109">CitationText</span></span>](citationtext.md)|
|[<span data-ttu-id="ca14e-110">Descrição</span><span class="sxs-lookup"><span data-stu-id="ca14e-110">Description</span></span>](description.md)|
|[<span data-ttu-id="ca14e-111">DictionaryName</span><span class="sxs-lookup"><span data-stu-id="ca14e-111">DictionaryName</span></span>](dictionaryname.md)|
|[<span data-ttu-id="ca14e-112">DictionaryHomePage</span><span class="sxs-lookup"><span data-stu-id="ca14e-112">DictionaryHomePage</span></span>](dictionaryhomepage.md)|
|[<span data-ttu-id="ca14e-113">DisplayName</span><span class="sxs-lookup"><span data-stu-id="ca14e-113">DisplayName</span></span>](displayname.md)|
|[<span data-ttu-id="ca14e-114">HighResolutionIconUrl</span><span class="sxs-lookup"><span data-stu-id="ca14e-114">HighResolutionIconUrl</span></span>](highresolutioniconurl.md)|
|[<span data-ttu-id="ca14e-115">IconUrl</span><span class="sxs-lookup"><span data-stu-id="ca14e-115">IconUrl</span></span>](iconurl.md)|
|[<span data-ttu-id="ca14e-116">QueryUri</span><span class="sxs-lookup"><span data-stu-id="ca14e-116">QueryUri</span></span>](queryuri.md)|
|[<span data-ttu-id="ca14e-117">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="ca14e-117">SourceLocation</span></span>](sourcelocation.md)|
|[<span data-ttu-id="ca14e-118">SupportUrl</span><span class="sxs-lookup"><span data-stu-id="ca14e-118">SupportUrl</span></span>](supporturl.md)|

## <a name="attributes"></a><span data-ttu-id="ca14e-119">Atributos</span><span class="sxs-lookup"><span data-stu-id="ca14e-119">Attributes</span></span>

|<span data-ttu-id="ca14e-120">Atributo</span><span class="sxs-lookup"><span data-stu-id="ca14e-120">Attribute</span></span>|<span data-ttu-id="ca14e-121">Tipo</span><span class="sxs-lookup"><span data-stu-id="ca14e-121">Type</span></span>|<span data-ttu-id="ca14e-122">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="ca14e-122">Required</span></span>|<span data-ttu-id="ca14e-123">Descrição</span><span class="sxs-lookup"><span data-stu-id="ca14e-123">Description</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="ca14e-124">Locale</span><span class="sxs-lookup"><span data-stu-id="ca14e-124">Locale</span></span>|<span data-ttu-id="ca14e-125">cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="ca14e-125">string</span></span>|<span data-ttu-id="ca14e-126">obrigatório</span><span class="sxs-lookup"><span data-stu-id="ca14e-126">required</span></span>|<span data-ttu-id="ca14e-127">Especifica o nome da cultura da localidade para essa substituição no formato de marca do idioma BCP 47, como `"en-US"`.</span><span class="sxs-lookup"><span data-stu-id="ca14e-127">Specifies the culture name of the locale for this override in the BCP 47 language tag format, such as  `"en-US"`.</span></span>|
|<span data-ttu-id="ca14e-128">Valor</span><span class="sxs-lookup"><span data-stu-id="ca14e-128">Value</span></span>|<span data-ttu-id="ca14e-129">cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="ca14e-129">string</span></span>|<span data-ttu-id="ca14e-130">obrigatório</span><span class="sxs-lookup"><span data-stu-id="ca14e-130">required</span></span>|<span data-ttu-id="ca14e-131">Especifica o valor da configuração expressa para a localidade especificada.</span><span class="sxs-lookup"><span data-stu-id="ca14e-131">Specifies value of the setting expressed for the specified locale.</span></span>|

## <a name="see-also"></a><span data-ttu-id="ca14e-132">Confira também</span><span class="sxs-lookup"><span data-stu-id="ca14e-132">See also</span></span>

- [<span data-ttu-id="ca14e-133">Localização para suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="ca14e-133">Localization for Office Add-ins</span></span>](../../develop/localization.md)
