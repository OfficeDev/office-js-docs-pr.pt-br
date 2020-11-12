---
title: Elemento ExtendedOverrides no arquivo de manifesto
description: Especifica as URLs para uma extensão formatada por JSON do manifesto.
ms.date: 11/06/2020
localization_priority: Normal
ms.openlocfilehash: 76491af34d1caf0ec266826df97a5363e336b85d
ms.sourcegitcommit: ca66ff7462bfdf4ed7ae04f43d1388c24de63bf9
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 11/11/2020
ms.locfileid: "48996670"
---
# <a name="extendedoverrides-element"></a><span data-ttu-id="f3504-103">Elemento ExtendedOverrides</span><span class="sxs-lookup"><span data-stu-id="f3504-103">ExtendedOverrides element</span></span>

<span data-ttu-id="f3504-104">Especifica as URLs completas para arquivos formatados por JSON que estendem o manifesto.</span><span class="sxs-lookup"><span data-stu-id="f3504-104">Specifies the full URLs for JSON-formatted files that extend the manifest.</span></span>

<span data-ttu-id="f3504-105">**Tipo de suplemento:** Painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="f3504-105">**Add-in type:** Task pane</span></span>

## <a name="syntax"></a><span data-ttu-id="f3504-106">Sintaxe</span><span class="sxs-lookup"><span data-stu-id="f3504-106">Syntax</span></span>

```XML
<ExtendedOverrides Url="string" [ResourcesUrl="string"] ></ExtendedOverrides>
```

## <a name="contained-in"></a><span data-ttu-id="f3504-107">Contido em</span><span class="sxs-lookup"><span data-stu-id="f3504-107">Contained in</span></span>

[<span data-ttu-id="f3504-108">OfficeApp</span><span class="sxs-lookup"><span data-stu-id="f3504-108">OfficeApp</span></span>](officeapp.md)

## <a name="can-contain"></a><span data-ttu-id="f3504-109">Pode conter</span><span class="sxs-lookup"><span data-stu-id="f3504-109">Can contain</span></span>

|<span data-ttu-id="f3504-110">Elemento</span><span class="sxs-lookup"><span data-stu-id="f3504-110">Element</span></span>|<span data-ttu-id="f3504-111">Conteúdo</span><span class="sxs-lookup"><span data-stu-id="f3504-111">Content</span></span>|<span data-ttu-id="f3504-112">Email</span><span class="sxs-lookup"><span data-stu-id="f3504-112">Mail</span></span>|<span data-ttu-id="f3504-113">TaskPane</span><span class="sxs-lookup"><span data-stu-id="f3504-113">TaskPane</span></span>|
|:-----|:-----|:-----|:-----|
|[<span data-ttu-id="f3504-114">Sinais</span><span class="sxs-lookup"><span data-stu-id="f3504-114">Tokens</span></span>](tokens.md)|||<span data-ttu-id="f3504-115">x</span><span class="sxs-lookup"><span data-stu-id="f3504-115">x</span></span>|

## <a name="attributes"></a><span data-ttu-id="f3504-116">Atributos</span><span class="sxs-lookup"><span data-stu-id="f3504-116">Attributes</span></span>

|<span data-ttu-id="f3504-117">Atributo</span><span class="sxs-lookup"><span data-stu-id="f3504-117">Attribute</span></span>|<span data-ttu-id="f3504-118">Descrição</span><span class="sxs-lookup"><span data-stu-id="f3504-118">Description</span></span>|
|:-----|:-----|
|<span data-ttu-id="f3504-119">URL (obrigatório)</span><span class="sxs-lookup"><span data-stu-id="f3504-119">Url (required)</span></span>| <span data-ttu-id="f3504-120">A URL completa do arquivo JSON de substituições estendidas.</span><span class="sxs-lookup"><span data-stu-id="f3504-120">The full URL of the extended overrides JSON file.</span></span> <span data-ttu-id="f3504-121">Pode ser um modelo de URL que usa tokens definidos pelo elemento [tokens](tokens.md) .</span><span class="sxs-lookup"><span data-stu-id="f3504-121">This could be a URL template that uses tokens defined by the [Tokens](tokens.md) element.</span></span>|
|<span data-ttu-id="f3504-122">ResourcesUrl (opcional)</span><span class="sxs-lookup"><span data-stu-id="f3504-122">ResourcesUrl (optional)</span></span> | <span data-ttu-id="f3504-123">A URL completa de um arquivo que fornece recursos suplementares, como cadeias de caracteres localizadas, para o arquivo especificado no `Url` atributo.</span><span class="sxs-lookup"><span data-stu-id="f3504-123">The full URL of a file that provides supplemental resources, such as localized strings, for the file specified in the `Url` attribute.</span></span> <span data-ttu-id="f3504-124">Pode ser um modelo de URL que usa tokens definidos pelo elemento [tokens](tokens.md) .</span><span class="sxs-lookup"><span data-stu-id="f3504-124">This could be a URL template that uses tokens defined by the [Tokens](tokens.md) element.</span></span>|

## <a name="example"></a><span data-ttu-id="f3504-125">Exemplo</span><span class="sxs-lookup"><span data-stu-id="f3504-125">Example</span></span>

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
