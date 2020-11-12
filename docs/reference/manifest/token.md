---
title: Elemento token no arquivo de manifesto
description: Especifica um token ou curinga que pode ser usado com modelos de URL no manifesto.
ms.date: 11/06/2020
localization_priority: Normal
ms.openlocfilehash: 5e26af44c566ab09ac81c8194e1ae7d85aaac327
ms.sourcegitcommit: ca66ff7462bfdf4ed7ae04f43d1388c24de63bf9
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 11/11/2020
ms.locfileid: "48996669"
---
# <a name="token-element"></a><span data-ttu-id="df823-103">Elemento token</span><span class="sxs-lookup"><span data-stu-id="df823-103">Token element</span></span>

<span data-ttu-id="df823-104">Define um token de URL individual.</span><span class="sxs-lookup"><span data-stu-id="df823-104">Defines an individual URL token.</span></span>

<span data-ttu-id="df823-105">**Tipo de suplemento:** Painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="df823-105">**Add-in type:** Task pane</span></span>

## <a name="syntax"></a><span data-ttu-id="df823-106">Sintaxe</span><span class="sxs-lookup"><span data-stu-id="df823-106">Syntax</span></span>

```XML
<Token Name="string" DefaultValue="string" xsi:type=["LocaleToken" | "RequirementsToken"] ></Token>
```

## <a name="contained-in"></a><span data-ttu-id="df823-107">Contido em</span><span class="sxs-lookup"><span data-stu-id="df823-107">Contained in</span></span>

[<span data-ttu-id="df823-108">Sinais</span><span class="sxs-lookup"><span data-stu-id="df823-108">Tokens</span></span>](tokens.md)

## <a name="can-contain"></a><span data-ttu-id="df823-109">Pode conter</span><span class="sxs-lookup"><span data-stu-id="df823-109">Can contain</span></span>

|<span data-ttu-id="df823-110">Elemento</span><span class="sxs-lookup"><span data-stu-id="df823-110">Element</span></span>|<span data-ttu-id="df823-111">Conteúdo</span><span class="sxs-lookup"><span data-stu-id="df823-111">Content</span></span>|<span data-ttu-id="df823-112">Email</span><span class="sxs-lookup"><span data-stu-id="df823-112">Mail</span></span>|<span data-ttu-id="df823-113">TaskPane</span><span class="sxs-lookup"><span data-stu-id="df823-113">TaskPane</span></span>|
|:-----|:-----|:-----|:-----|
|[<span data-ttu-id="df823-114">Override</span><span class="sxs-lookup"><span data-stu-id="df823-114">Override</span></span>](override.md)|||<span data-ttu-id="df823-115">x</span><span class="sxs-lookup"><span data-stu-id="df823-115">x</span></span>|

## <a name="attributes"></a><span data-ttu-id="df823-116">Atributos</span><span class="sxs-lookup"><span data-stu-id="df823-116">Attributes</span></span>

|<span data-ttu-id="df823-117">Atributo</span><span class="sxs-lookup"><span data-stu-id="df823-117">Attribute</span></span>|<span data-ttu-id="df823-118">Descrição</span><span class="sxs-lookup"><span data-stu-id="df823-118">Description</span></span>|
|:-----|:-----|
|<span data-ttu-id="df823-119">DefaultValue</span><span class="sxs-lookup"><span data-stu-id="df823-119">DefaultValue</span></span>|<span data-ttu-id="df823-120">Valor padrão para esse token se nenhuma condição em qualquer `<Override>` elemento filho corresponder.</span><span class="sxs-lookup"><span data-stu-id="df823-120">Default value for this token if no condition in any child `<Override>` element matches.</span></span>|
|<span data-ttu-id="df823-121">Nome</span><span class="sxs-lookup"><span data-stu-id="df823-121">Name</span></span>|<span data-ttu-id="df823-122">Nome do token.</span><span class="sxs-lookup"><span data-stu-id="df823-122">Token name.</span></span> <span data-ttu-id="df823-123">Esse nome é definido pelo usuário.</span><span class="sxs-lookup"><span data-stu-id="df823-123">This name is user-defined.</span></span> <span data-ttu-id="df823-124">O tipo do token é determinado pelo atributo Type.</span><span class="sxs-lookup"><span data-stu-id="df823-124">The type of the token is determined by the type attribute.</span></span>|
|<span data-ttu-id="df823-125">xsi:type</span><span class="sxs-lookup"><span data-stu-id="df823-125">xsi:type</span></span>|<span data-ttu-id="df823-126">Define o tipo de token.</span><span class="sxs-lookup"><span data-stu-id="df823-126">Defines the kind of Token.</span></span> <span data-ttu-id="df823-127">Este atributo deve ser definido como um de:  `"RequirementsToken"` ou  `"LocaleToken"` .</span><span class="sxs-lookup"><span data-stu-id="df823-127">This attribute should be set to one of:  `"RequirementsToken"`,  or  `"LocaleToken"`.</span></span>|

## <a name="example"></a><span data-ttu-id="df823-128">Exemplo</span><span class="sxs-lookup"><span data-stu-id="df823-128">Example</span></span>

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