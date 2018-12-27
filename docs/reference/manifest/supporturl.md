---
title: Elemento SupportUrl no arquivo de manifesto
description: ''
ms.date: 10/09/2018
ms.openlocfilehash: 00234ef9fe8960b9956e6a2595e2e2e71bfb97c6
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 12/22/2018
ms.locfileid: "27432666"
---
# <a name="supporturl-element"></a><span data-ttu-id="56fa3-102">Elemento SupportUrl</span><span class="sxs-lookup"><span data-stu-id="56fa3-102">SupportUrl element</span></span>

<span data-ttu-id="56fa3-103">Especifica a URL de uma página que fornece informações de suporte para o suplemento.</span><span class="sxs-lookup"><span data-stu-id="56fa3-103">Specifies the URL of a page that provides support information for your add-in.</span></span>

## <a name="syntax"></a><span data-ttu-id="56fa3-104">Sintaxe</span><span class="sxs-lookup"><span data-stu-id="56fa3-104">Syntax</span></span>

```XML
<OfficeApp>
...
  <IconUrl DefaultValue="https://contoso.com/assets/icon-32.png" />
  <HighResolutionIconUrl DefaultValue="https://contoso.com/assets/hi-res-icon.png"/>
  
  
  <SupportUrl DefaultValue="https://contoso.com/support " />
  
  
  <AppDomains>
  ...
  </AppDomains>
...
</OfficeApp>
```

## <a name="contained-in"></a><span data-ttu-id="56fa3-105">Contido em</span><span class="sxs-lookup"><span data-stu-id="56fa3-105">Contained in</span></span>

[<span data-ttu-id="56fa3-106">OfficeApp</span><span class="sxs-lookup"><span data-stu-id="56fa3-106">OfficeApp</span></span>](officeapp.md)

## <a name="can-contain"></a><span data-ttu-id="56fa3-107">Pode conter</span><span class="sxs-lookup"><span data-stu-id="56fa3-107">Can contain</span></span>

|  <span data-ttu-id="56fa3-108">Elemento</span><span class="sxs-lookup"><span data-stu-id="56fa3-108">Element</span></span> | <span data-ttu-id="56fa3-109">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="56fa3-109">Required</span></span> | <span data-ttu-id="56fa3-110">Descrição</span><span class="sxs-lookup"><span data-stu-id="56fa3-110">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="56fa3-111">Override</span><span class="sxs-lookup"><span data-stu-id="56fa3-111">Override</span></span>](override.md)   | <span data-ttu-id="56fa3-112">Não</span><span class="sxs-lookup"><span data-stu-id="56fa3-112">No</span></span> | <span data-ttu-id="56fa3-113">Especifica a configuração de URLs de localidades adicionais</span><span class="sxs-lookup"><span data-stu-id="56fa3-113">Specifies the setting for additional locale urls</span></span> |

## <a name="attributes"></a><span data-ttu-id="56fa3-114">Atributos</span><span class="sxs-lookup"><span data-stu-id="56fa3-114">Attributes</span></span>

|<span data-ttu-id="56fa3-115">**Atributo**</span><span class="sxs-lookup"><span data-stu-id="56fa3-115">**Attribute**</span></span>|<span data-ttu-id="56fa3-116">**Tipo**</span><span class="sxs-lookup"><span data-stu-id="56fa3-116">**Type**</span></span>|<span data-ttu-id="56fa3-117">**Obrigatório**</span><span class="sxs-lookup"><span data-stu-id="56fa3-117">**Required**</span></span>|<span data-ttu-id="56fa3-118">**Descrição**</span><span class="sxs-lookup"><span data-stu-id="56fa3-118">**Description**</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="56fa3-119">DefaultValue</span><span class="sxs-lookup"><span data-stu-id="56fa3-119">DefaultValue</span></span>|<span data-ttu-id="56fa3-120">URL</span><span class="sxs-lookup"><span data-stu-id="56fa3-120">URL</span></span>|<span data-ttu-id="56fa3-121">obrigatório</span><span class="sxs-lookup"><span data-stu-id="56fa3-121">required</span></span>|<span data-ttu-id="56fa3-122">Especifica o valor padrão para essa configuração, expresso para a localidade especificada no elemento [DefaultLocale](defaultlocale.md).</span><span class="sxs-lookup"><span data-stu-id="56fa3-122">Specifies the default value for this setting, expressed for the locale specified in the [DefaultLocale](defaultlocale.md) element.</span></span>|
