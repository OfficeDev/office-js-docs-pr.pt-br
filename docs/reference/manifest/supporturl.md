---
title: Elemento SupportUrl no arquivo de manifesto
description: O elemento SupportUrl especifica a URL de uma página que fornece informações de suporte para o suplemento.
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: e38030062c48936f925126e896cd74e660164a5d
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/17/2020
ms.locfileid: "42720341"
---
# <a name="supporturl-element"></a><span data-ttu-id="7d01c-103">Elemento SupportUrl</span><span class="sxs-lookup"><span data-stu-id="7d01c-103">SupportUrl element</span></span>

<span data-ttu-id="7d01c-104">Especifica a URL de uma página que fornece informações de suporte para o suplemento.</span><span class="sxs-lookup"><span data-stu-id="7d01c-104">Specifies the URL of a page that provides support information for your add-in.</span></span>

## <a name="syntax"></a><span data-ttu-id="7d01c-105">Sintaxe</span><span class="sxs-lookup"><span data-stu-id="7d01c-105">Syntax</span></span>

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

## <a name="contained-in"></a><span data-ttu-id="7d01c-106">Contido em</span><span class="sxs-lookup"><span data-stu-id="7d01c-106">Contained in</span></span>

[<span data-ttu-id="7d01c-107">OfficeApp</span><span class="sxs-lookup"><span data-stu-id="7d01c-107">OfficeApp</span></span>](officeapp.md)

## <a name="can-contain"></a><span data-ttu-id="7d01c-108">Pode conter</span><span class="sxs-lookup"><span data-stu-id="7d01c-108">Can contain</span></span>

|  <span data-ttu-id="7d01c-109">Elemento</span><span class="sxs-lookup"><span data-stu-id="7d01c-109">Element</span></span> | <span data-ttu-id="7d01c-110">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="7d01c-110">Required</span></span> | <span data-ttu-id="7d01c-111">Descrição</span><span class="sxs-lookup"><span data-stu-id="7d01c-111">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="7d01c-112">Override</span><span class="sxs-lookup"><span data-stu-id="7d01c-112">Override</span></span>](override.md)   | <span data-ttu-id="7d01c-113">Não</span><span class="sxs-lookup"><span data-stu-id="7d01c-113">No</span></span> | <span data-ttu-id="7d01c-114">Especifica a configuração de URLs de localidades adicionais</span><span class="sxs-lookup"><span data-stu-id="7d01c-114">Specifies the setting for additional locale urls</span></span> |

## <a name="attributes"></a><span data-ttu-id="7d01c-115">Atributos</span><span class="sxs-lookup"><span data-stu-id="7d01c-115">Attributes</span></span>

|<span data-ttu-id="7d01c-116">**Atributo**</span><span class="sxs-lookup"><span data-stu-id="7d01c-116">**Attribute**</span></span>|<span data-ttu-id="7d01c-117">**Tipo**</span><span class="sxs-lookup"><span data-stu-id="7d01c-117">**Type**</span></span>|<span data-ttu-id="7d01c-118">**Obrigatório**</span><span class="sxs-lookup"><span data-stu-id="7d01c-118">**Required**</span></span>|<span data-ttu-id="7d01c-119">**Descrição**</span><span class="sxs-lookup"><span data-stu-id="7d01c-119">**Description**</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="7d01c-120">DefaultValue</span><span class="sxs-lookup"><span data-stu-id="7d01c-120">DefaultValue</span></span>|<span data-ttu-id="7d01c-121">URL</span><span class="sxs-lookup"><span data-stu-id="7d01c-121">URL</span></span>|<span data-ttu-id="7d01c-122">obrigatório</span><span class="sxs-lookup"><span data-stu-id="7d01c-122">required</span></span>|<span data-ttu-id="7d01c-123">Especifica o valor padrão para essa configuração, expresso para a localidade especificada no elemento [DefaultLocale](defaultlocale.md).</span><span class="sxs-lookup"><span data-stu-id="7d01c-123">Specifies the default value for this setting, expressed for the locale specified in the [DefaultLocale](defaultlocale.md) element.</span></span>|
