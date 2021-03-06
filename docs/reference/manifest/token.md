---
title: Elemento Token no arquivo de manifesto
description: Especifica um token ou curinga que pode ser usado com modelos de URL no manifesto.
ms.date: 11/06/2020
localization_priority: Normal
ms.openlocfilehash: 48078f8211a8fd3f0e3f9d7c3f3aabd1d31b0a6d
ms.sourcegitcommit: e7009c565b18c607fe0868db2e26e250ad308dce
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/05/2021
ms.locfileid: "50505364"
---
# <a name="token-element"></a><span data-ttu-id="20069-103">Elemento Token</span><span class="sxs-lookup"><span data-stu-id="20069-103">Token element</span></span>

<span data-ttu-id="20069-104">Define um token de URL individual.</span><span class="sxs-lookup"><span data-stu-id="20069-104">Defines an individual URL token.</span></span> <span data-ttu-id="20069-105">Para obter mais informações sobre o uso desse elemento, consulte [Trabalhar com substituições estendidas do manifesto](../../develop/extended-overrides.md).</span><span class="sxs-lookup"><span data-stu-id="20069-105">For more information about the use of this element, see [Work with extended overrides of the manifest](../../develop/extended-overrides.md).</span></span>

<span data-ttu-id="20069-106">**Tipo de suplemento:** Painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="20069-106">**Add-in type:** Task pane</span></span>

## <a name="syntax"></a><span data-ttu-id="20069-107">Sintaxe</span><span class="sxs-lookup"><span data-stu-id="20069-107">Syntax</span></span>

```XML
<Token Name="string" DefaultValue="string" xsi:type=["LocaleToken" | "RequirementsToken"] ></Token>
```

## <a name="contained-in"></a><span data-ttu-id="20069-108">Contido em</span><span class="sxs-lookup"><span data-stu-id="20069-108">Contained in</span></span>

[<span data-ttu-id="20069-109">Tokens</span><span class="sxs-lookup"><span data-stu-id="20069-109">Tokens</span></span>](tokens.md)

## <a name="can-contain"></a><span data-ttu-id="20069-110">Pode conter</span><span class="sxs-lookup"><span data-stu-id="20069-110">Can contain</span></span>

|<span data-ttu-id="20069-111">Elemento</span><span class="sxs-lookup"><span data-stu-id="20069-111">Element</span></span>|<span data-ttu-id="20069-112">Conteúdo</span><span class="sxs-lookup"><span data-stu-id="20069-112">Content</span></span>|<span data-ttu-id="20069-113">Email</span><span class="sxs-lookup"><span data-stu-id="20069-113">Mail</span></span>|<span data-ttu-id="20069-114">TaskPane</span><span class="sxs-lookup"><span data-stu-id="20069-114">TaskPane</span></span>|
|:-----|:-----|:-----|:-----|
|[<span data-ttu-id="20069-115">Override</span><span class="sxs-lookup"><span data-stu-id="20069-115">Override</span></span>](override.md)|||<span data-ttu-id="20069-116">x</span><span class="sxs-lookup"><span data-stu-id="20069-116">x</span></span>|

## <a name="attributes"></a><span data-ttu-id="20069-117">Atributos</span><span class="sxs-lookup"><span data-stu-id="20069-117">Attributes</span></span>

|<span data-ttu-id="20069-118">Atributo</span><span class="sxs-lookup"><span data-stu-id="20069-118">Attribute</span></span>|<span data-ttu-id="20069-119">Descrição</span><span class="sxs-lookup"><span data-stu-id="20069-119">Description</span></span>|
|:-----|:-----|
|<span data-ttu-id="20069-120">DefaultValue</span><span class="sxs-lookup"><span data-stu-id="20069-120">DefaultValue</span></span>|<span data-ttu-id="20069-121">Valor padrão para esse token se nenhuma condição em qualquer `<Override>` elemento filho corresponde.</span><span class="sxs-lookup"><span data-stu-id="20069-121">Default value for this token if no condition in any child `<Override>` element matches.</span></span>|
|<span data-ttu-id="20069-122">Nome</span><span class="sxs-lookup"><span data-stu-id="20069-122">Name</span></span>|<span data-ttu-id="20069-123">Nome do token.</span><span class="sxs-lookup"><span data-stu-id="20069-123">Token name.</span></span> <span data-ttu-id="20069-124">Esse nome é definido pelo usuário.</span><span class="sxs-lookup"><span data-stu-id="20069-124">This name is user-defined.</span></span> <span data-ttu-id="20069-125">O tipo do token é determinado pelo atributo type.</span><span class="sxs-lookup"><span data-stu-id="20069-125">The type of the token is determined by the type attribute.</span></span>|
|<span data-ttu-id="20069-126">xsi:type</span><span class="sxs-lookup"><span data-stu-id="20069-126">xsi:type</span></span>|<span data-ttu-id="20069-127">Define o tipo de Token.</span><span class="sxs-lookup"><span data-stu-id="20069-127">Defines the kind of Token.</span></span> <span data-ttu-id="20069-128">Esse atributo deve ser definido como um dos:  `"RequirementsToken"` , ou  `"LocaleToken"` .</span><span class="sxs-lookup"><span data-stu-id="20069-128">This attribute should be set to one of:  `"RequirementsToken"`,  or  `"LocaleToken"`.</span></span>|

## <a name="example"></a><span data-ttu-id="20069-129">Exemplo</span><span class="sxs-lookup"><span data-stu-id="20069-129">Example</span></span>

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