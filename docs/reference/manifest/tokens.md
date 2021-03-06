---
title: Elemento Tokens no arquivo de manifesto
description: Especifica tokens ou curingas que podem ser usados com modelos de URL no manifesto.
ms.date: 11/06/2020
localization_priority: Normal
ms.openlocfilehash: 8680b985068c44e93f601a2b24e2f28899eb483d
ms.sourcegitcommit: e7009c565b18c607fe0868db2e26e250ad308dce
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/05/2021
ms.locfileid: "50505322"
---
# <a name="tokens-element"></a><span data-ttu-id="35f47-103">Elemento Tokens</span><span class="sxs-lookup"><span data-stu-id="35f47-103">Tokens element</span></span>

<span data-ttu-id="35f47-104">Define tokens que podem ser usados em URLs de modelo.</span><span class="sxs-lookup"><span data-stu-id="35f47-104">Defines tokens that could be used in template URLs.</span></span> <span data-ttu-id="35f47-105">Para obter mais informações sobre o uso desse elemento, consulte [Trabalhar com substituições estendidas do manifesto](../../develop/extended-overrides.md).</span><span class="sxs-lookup"><span data-stu-id="35f47-105">For more information about the use of this element, see [Work with extended overrides of the manifest](../../develop/extended-overrides.md).</span></span>

<span data-ttu-id="35f47-106">**Tipo de suplemento:** Painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="35f47-106">**Add-in type:** Task pane</span></span>

## <a name="syntax"></a><span data-ttu-id="35f47-107">Sintaxe</span><span class="sxs-lookup"><span data-stu-id="35f47-107">Syntax</span></span>

```XML
<Tokens></Tokens>
```

## <a name="contained-in"></a><span data-ttu-id="35f47-108">Contido em</span><span class="sxs-lookup"><span data-stu-id="35f47-108">Contained in</span></span>

[<span data-ttu-id="35f47-109">ExtendedOverrides</span><span class="sxs-lookup"><span data-stu-id="35f47-109">ExtendedOverrides</span></span>](extendedoverrides.md)

## <a name="must-contain"></a><span data-ttu-id="35f47-110">Deve conter</span><span class="sxs-lookup"><span data-stu-id="35f47-110">Must contain</span></span>

|<span data-ttu-id="35f47-111">Elemento</span><span class="sxs-lookup"><span data-stu-id="35f47-111">Element</span></span>|<span data-ttu-id="35f47-112">Conteúdo</span><span class="sxs-lookup"><span data-stu-id="35f47-112">Content</span></span>|<span data-ttu-id="35f47-113">Email</span><span class="sxs-lookup"><span data-stu-id="35f47-113">Mail</span></span>|<span data-ttu-id="35f47-114">TaskPane</span><span class="sxs-lookup"><span data-stu-id="35f47-114">TaskPane</span></span>|
|:-----|:-----|:-----|:-----|
|[<span data-ttu-id="35f47-115">Token</span><span class="sxs-lookup"><span data-stu-id="35f47-115">Token</span></span>](token.md)|||<span data-ttu-id="35f47-116">x</span><span class="sxs-lookup"><span data-stu-id="35f47-116">x</span></span>|

## <a name="example"></a><span data-ttu-id="35f47-117">Exemplo</span><span class="sxs-lookup"><span data-stu-id="35f47-117">Example</span></span>

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