---
title: Elemento Override no arquivo de manifesto
description: ''
ms.date: 10/09/2018
ms.openlocfilehash: d1d2400312f12116b1ac5f4010135541e783dcc7
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 12/22/2018
ms.locfileid: "27432862"
---
# <a name="override-element"></a><span data-ttu-id="34df2-102">Elemento Override</span><span class="sxs-lookup"><span data-stu-id="34df2-102">Override element</span></span>

<span data-ttu-id="34df2-103">Fornece uma maneira de especificar o valor de uma configuração para uma localidade adicional.</span><span class="sxs-lookup"><span data-stu-id="34df2-103">Provides a way to specify the value of a setting for an additional locale.</span></span>

<span data-ttu-id="34df2-104">**Tipo de suplemento:** Conteúdo, Painel de tarefas, Email</span><span class="sxs-lookup"><span data-stu-id="34df2-104">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="34df2-105">Sintaxe</span><span class="sxs-lookup"><span data-stu-id="34df2-105">Syntax</span></span>

```XML
<Override Locale="string" Value="string" />
```

## <a name="contained-in"></a><span data-ttu-id="34df2-106">Contido em</span><span class="sxs-lookup"><span data-stu-id="34df2-106">Contained in</span></span>

|<span data-ttu-id="34df2-107">**Elemento**</span><span class="sxs-lookup"><span data-stu-id="34df2-107">**Element**</span></span>|
|:-----|
|[<span data-ttu-id="34df2-108">CitationText</span><span class="sxs-lookup"><span data-stu-id="34df2-108">CitationText</span></span>](citationtext.md)|
|[<span data-ttu-id="34df2-109">Description</span><span class="sxs-lookup"><span data-stu-id="34df2-109">Description</span></span>](description.md)|
|[<span data-ttu-id="34df2-110">DictionaryName</span><span class="sxs-lookup"><span data-stu-id="34df2-110">DictionaryName</span></span>](dictionaryname.md)|
|[<span data-ttu-id="34df2-111">DictionaryHomePage</span><span class="sxs-lookup"><span data-stu-id="34df2-111">DictionaryHomePage</span></span>](dictionaryhomepage.md)|
|[<span data-ttu-id="34df2-112">DisplayName</span><span class="sxs-lookup"><span data-stu-id="34df2-112">DisplayName</span></span>](displayname.md)|
|[<span data-ttu-id="34df2-113">HighResolutionIconUrl</span><span class="sxs-lookup"><span data-stu-id="34df2-113">HighResolutionIconUrl</span></span>](highresolutioniconurl.md)|
|[<span data-ttu-id="34df2-114">IconUrl</span><span class="sxs-lookup"><span data-stu-id="34df2-114">IconUrl</span></span>](iconurl.md)|
|[<span data-ttu-id="34df2-115">QueryUri</span><span class="sxs-lookup"><span data-stu-id="34df2-115">QueryUri</span></span>](queryuri.md)|
|[<span data-ttu-id="34df2-116">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="34df2-116">SourceLocation</span></span>](sourcelocation.md)|
|[<span data-ttu-id="34df2-117">SupportUrl</span><span class="sxs-lookup"><span data-stu-id="34df2-117">SupportUrl</span></span>](supporturl.md)|

## <a name="attributes"></a><span data-ttu-id="34df2-118">Atributos</span><span class="sxs-lookup"><span data-stu-id="34df2-118">Attributes</span></span>

|<span data-ttu-id="34df2-119">**Atributo**</span><span class="sxs-lookup"><span data-stu-id="34df2-119">**Attribute**</span></span>|<span data-ttu-id="34df2-120">**Tipo**</span><span class="sxs-lookup"><span data-stu-id="34df2-120">**Type**</span></span>|<span data-ttu-id="34df2-121">**Obrigatório**</span><span class="sxs-lookup"><span data-stu-id="34df2-121">**Required**</span></span>|<span data-ttu-id="34df2-122">**Descrição**</span><span class="sxs-lookup"><span data-stu-id="34df2-122">**Description**</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="34df2-123">Locale</span><span class="sxs-lookup"><span data-stu-id="34df2-123">Locale</span></span>|<span data-ttu-id="34df2-124">cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="34df2-124">string</span></span>|<span data-ttu-id="34df2-125">obrigatório</span><span class="sxs-lookup"><span data-stu-id="34df2-125">required</span></span>|<span data-ttu-id="34df2-126">Especifica o nome da cultura da localidade para essa substituição no formato de marca do idioma BCP 47, como `"en-US"`.</span><span class="sxs-lookup"><span data-stu-id="34df2-126">Specifies the culture name of the locale for this override in the BCP 47 language tag format, such as  `"en-US"`.</span></span>|
|<span data-ttu-id="34df2-127">Valor</span><span class="sxs-lookup"><span data-stu-id="34df2-127">Value</span></span>|<span data-ttu-id="34df2-128">cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="34df2-128">string</span></span>|<span data-ttu-id="34df2-129">obrigatório</span><span class="sxs-lookup"><span data-stu-id="34df2-129">required</span></span>|<span data-ttu-id="34df2-130">Especifica o valor da configuração expressa para a localidade especificada.</span><span class="sxs-lookup"><span data-stu-id="34df2-130">Specifies value of the setting expressed for the specified locale.</span></span>|

## <a name="see-also"></a><span data-ttu-id="34df2-131">Confira também</span><span class="sxs-lookup"><span data-stu-id="34df2-131">See also</span></span>

- [<span data-ttu-id="34df2-132">Localização para suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="34df2-132">Localization for Office Add-ins</span></span>](https://docs.microsoft.com/office/dev/add-ins/develop/localization)
    
