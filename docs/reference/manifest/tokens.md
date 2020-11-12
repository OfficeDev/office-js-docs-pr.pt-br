---
title: Elemento tokens no arquivo de manifesto
description: Especifica tokens ou curingas que podem ser usados com modelos de URL no manifesto.
ms.date: 11/06/2020
localization_priority: Normal
ms.openlocfilehash: a50de7c2c3e8ebeb9425c1677a94bbcc62281d3b
ms.sourcegitcommit: ca66ff7462bfdf4ed7ae04f43d1388c24de63bf9
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 11/11/2020
ms.locfileid: "48996667"
---
# <a name="tokens-element"></a><span data-ttu-id="54af7-103">Elemento tokens</span><span class="sxs-lookup"><span data-stu-id="54af7-103">Tokens element</span></span>

<span data-ttu-id="54af7-104">Define tokens que podem ser usados em URLs de modelo.</span><span class="sxs-lookup"><span data-stu-id="54af7-104">Defines tokens that could be used in template URLs.</span></span>

<span data-ttu-id="54af7-105">**Tipo de suplemento:** Painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="54af7-105">**Add-in type:** Task pane</span></span>

## <a name="syntax"></a><span data-ttu-id="54af7-106">Sintaxe</span><span class="sxs-lookup"><span data-stu-id="54af7-106">Syntax</span></span>

```XML
<Tokens></Tokens>
```

## <a name="contained-in"></a><span data-ttu-id="54af7-107">Contido em</span><span class="sxs-lookup"><span data-stu-id="54af7-107">Contained in</span></span>

[<span data-ttu-id="54af7-108">ExtendedOverrides</span><span class="sxs-lookup"><span data-stu-id="54af7-108">ExtendedOverrides</span></span>](extendedoverrides.md)

## <a name="must-contain"></a><span data-ttu-id="54af7-109">Deve conter</span><span class="sxs-lookup"><span data-stu-id="54af7-109">Must contain</span></span>

|<span data-ttu-id="54af7-110">Elemento</span><span class="sxs-lookup"><span data-stu-id="54af7-110">Element</span></span>|<span data-ttu-id="54af7-111">Conteúdo</span><span class="sxs-lookup"><span data-stu-id="54af7-111">Content</span></span>|<span data-ttu-id="54af7-112">Email</span><span class="sxs-lookup"><span data-stu-id="54af7-112">Mail</span></span>|<span data-ttu-id="54af7-113">TaskPane</span><span class="sxs-lookup"><span data-stu-id="54af7-113">TaskPane</span></span>|
|:-----|:-----|:-----|:-----|
|[<span data-ttu-id="54af7-114">Token</span><span class="sxs-lookup"><span data-stu-id="54af7-114">Token</span></span>](token.md)|||<span data-ttu-id="54af7-115">x</span><span class="sxs-lookup"><span data-stu-id="54af7-115">x</span></span>|

## <a name="example"></a><span data-ttu-id="54af7-116">Exemplo</span><span class="sxs-lookup"><span data-stu-id="54af7-116">Example</span></span>

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