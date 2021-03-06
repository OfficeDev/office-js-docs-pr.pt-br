---
title: Elemento ExtendedOverrides no arquivo de manifesto
description: Especifica as URLs para uma extensão formatada por JSON do manifesto.
ms.date: 02/23/2021
localization_priority: Normal
ms.openlocfilehash: f433c9c5604f3fae35580ba20780ea6fe91401c7
ms.sourcegitcommit: e7009c565b18c607fe0868db2e26e250ad308dce
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/05/2021
ms.locfileid: "50505469"
---
# <a name="extendedoverrides-element"></a><span data-ttu-id="2b91d-103">Elemento ExtendedOverrides</span><span class="sxs-lookup"><span data-stu-id="2b91d-103">ExtendedOverrides element</span></span>

<span data-ttu-id="2b91d-104">Especifica as URLs completas para arquivos formatados com JSON que estendem o manifesto.</span><span class="sxs-lookup"><span data-stu-id="2b91d-104">Specifies the full URLs for JSON-formatted files that extend the manifest.</span></span> <span data-ttu-id="2b91d-105">Para obter informações detalhadas sobre o uso desse elemento e seus elementos [descendentes,](../../develop/extended-overrides.md)consulte Trabalhar com substituições estendidas do manifesto .</span><span class="sxs-lookup"><span data-stu-id="2b91d-105">For detailed information about the use of this element and its descendent elements, see [Work with extended overrides of the manifest](../../develop/extended-overrides.md).</span></span>

<span data-ttu-id="2b91d-106">**Tipo de suplemento:** Painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="2b91d-106">**Add-in type:** Task pane</span></span>

## <a name="syntax"></a><span data-ttu-id="2b91d-107">Sintaxe</span><span class="sxs-lookup"><span data-stu-id="2b91d-107">Syntax</span></span>

```XML
<ExtendedOverrides Url="string" [ResourcesUrl="string"] ></ExtendedOverrides>
```

## <a name="contained-in"></a><span data-ttu-id="2b91d-108">Contido em</span><span class="sxs-lookup"><span data-stu-id="2b91d-108">Contained in</span></span>

[<span data-ttu-id="2b91d-109">OfficeApp</span><span class="sxs-lookup"><span data-stu-id="2b91d-109">OfficeApp</span></span>](officeapp.md)

## <a name="can-contain"></a><span data-ttu-id="2b91d-110">Pode conter</span><span class="sxs-lookup"><span data-stu-id="2b91d-110">Can contain</span></span>

|<span data-ttu-id="2b91d-111">Elemento</span><span class="sxs-lookup"><span data-stu-id="2b91d-111">Element</span></span>|<span data-ttu-id="2b91d-112">Conteúdo</span><span class="sxs-lookup"><span data-stu-id="2b91d-112">Content</span></span>|<span data-ttu-id="2b91d-113">Email</span><span class="sxs-lookup"><span data-stu-id="2b91d-113">Mail</span></span>|<span data-ttu-id="2b91d-114">TaskPane</span><span class="sxs-lookup"><span data-stu-id="2b91d-114">TaskPane</span></span>|
|:-----|:-----|:-----|:-----|
|[<span data-ttu-id="2b91d-115">Tokens</span><span class="sxs-lookup"><span data-stu-id="2b91d-115">Tokens</span></span>](tokens.md)|||<span data-ttu-id="2b91d-116">x</span><span class="sxs-lookup"><span data-stu-id="2b91d-116">x</span></span>|

## <a name="attributes"></a><span data-ttu-id="2b91d-117">Atributos</span><span class="sxs-lookup"><span data-stu-id="2b91d-117">Attributes</span></span>

|<span data-ttu-id="2b91d-118">Atributo</span><span class="sxs-lookup"><span data-stu-id="2b91d-118">Attribute</span></span>|<span data-ttu-id="2b91d-119">Descrição</span><span class="sxs-lookup"><span data-stu-id="2b91d-119">Description</span></span>|
|:-----|:-----|
|<span data-ttu-id="2b91d-120">URL (obrigatório)</span><span class="sxs-lookup"><span data-stu-id="2b91d-120">Url (required)</span></span>| <span data-ttu-id="2b91d-121">A URL completa do arquivo JSON substitui estendido.</span><span class="sxs-lookup"><span data-stu-id="2b91d-121">The full URL of the extended overrides JSON file.</span></span> <span data-ttu-id="2b91d-122">No futuro, esse valor pode ser um modelo de URL que usa tokens definidos pelo [elemento Tokens.](tokens.md)</span><span class="sxs-lookup"><span data-stu-id="2b91d-122">In the future, this value could be a URL template that uses tokens defined by the [Tokens](tokens.md) element.</span></span> <span data-ttu-id="2b91d-123">Consulte [Exemplos](#examples).</span><span class="sxs-lookup"><span data-stu-id="2b91d-123">See [Examples](#examples).</span></span>|
|<span data-ttu-id="2b91d-124">ResourcesUrl (opcional)</span><span class="sxs-lookup"><span data-stu-id="2b91d-124">ResourcesUrl (optional)</span></span> | <span data-ttu-id="2b91d-125">A URL completa de um arquivo que fornece recursos suplementares, como cadeias de caracteres localizadas, para o arquivo especificado no `Url` atributo.</span><span class="sxs-lookup"><span data-stu-id="2b91d-125">The full URL of a file that provides supplemental resources, such as localized strings, for the file specified in the `Url` attribute.</span></span> <span data-ttu-id="2b91d-126">Pode ser um modelo de URL que usa tokens definidos pelo [elemento Tokens.](tokens.md)</span><span class="sxs-lookup"><span data-stu-id="2b91d-126">This could be a URL template that uses tokens defined by the [Tokens](tokens.md) element.</span></span>|

## <a name="examples"></a><span data-ttu-id="2b91d-127">Exemplos</span><span class="sxs-lookup"><span data-stu-id="2b91d-127">Examples</span></span>

```XML
<OfficeApp ...>
  <!-- other elements omitted -->
  <ExtendedOverrides Url="http://contoso.com/addinmetadata/extended-manifest-overrides.json"
                     ResourceUrl="https://contoso.com/addin/my-resources.json">
  </ExtendedOverrides>
</OfficeApp>
```

<span data-ttu-id="2b91d-128">No futuro, esse valor pode ser um modelo de URL que usa tokens definidos pelo [elemento Tokens.](tokens.md)</span><span class="sxs-lookup"><span data-stu-id="2b91d-128">In the future, this value could be a URL template that uses tokens defined by the [Tokens](tokens.md) element.</span></span> <span data-ttu-id="2b91d-129">Apresentamos um exemplo a seguir.</span><span class="sxs-lookup"><span data-stu-id="2b91d-129">The following is an example.</span></span>

```XML
<OfficeApp ...>
  <!-- other elements omitted -->
  <ExtendedOverrides Url="http://contoso.com/addinmetadata/${token.locale}/extended-manifest-overrides.json">
    <Tokens>
      <Token Name="locale" DefaultValue="en-us" xsi:type="LocaleToken">
        <Override Locale="es-*" Value="es-es" />
        <Override Locale="es-mx" Value="es-mx" />
        <Override Locale="fr-*" Value="fr-fr" />
        <Override Locale="ja-jp" Value="ja-jp" />
      </Token>
    <Tokens>
  </ExtendedOverrides>
</OfficeApp>
```
