---
title: Elemento SupportUrl no arquivo de manifesto
description: O elemento SupportUrl especifica a URL de uma página que fornece informações de suporte para o suplemento.
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: f75ee811699823a501ac594e66daaaf3f93c2782
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 06/08/2020
ms.locfileid: "44608702"
---
# <a name="supporturl-element"></a><span data-ttu-id="0da3b-103">Elemento SupportUrl</span><span class="sxs-lookup"><span data-stu-id="0da3b-103">SupportUrl element</span></span>

<span data-ttu-id="0da3b-104">Especifica a URL de uma página que fornece informações de suporte para o suplemento.</span><span class="sxs-lookup"><span data-stu-id="0da3b-104">Specifies the URL of a page that provides support information for your add-in.</span></span>

## <a name="syntax"></a><span data-ttu-id="0da3b-105">Sintaxe</span><span class="sxs-lookup"><span data-stu-id="0da3b-105">Syntax</span></span>

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

## <a name="contained-in"></a><span data-ttu-id="0da3b-106">Contido em</span><span class="sxs-lookup"><span data-stu-id="0da3b-106">Contained in</span></span>

[<span data-ttu-id="0da3b-107">OfficeApp</span><span class="sxs-lookup"><span data-stu-id="0da3b-107">OfficeApp</span></span>](officeapp.md)

## <a name="can-contain"></a><span data-ttu-id="0da3b-108">Pode conter</span><span class="sxs-lookup"><span data-stu-id="0da3b-108">Can contain</span></span>

|  <span data-ttu-id="0da3b-109">Elemento</span><span class="sxs-lookup"><span data-stu-id="0da3b-109">Element</span></span> | <span data-ttu-id="0da3b-110">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="0da3b-110">Required</span></span> | <span data-ttu-id="0da3b-111">Descrição</span><span class="sxs-lookup"><span data-stu-id="0da3b-111">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="0da3b-112">Override</span><span class="sxs-lookup"><span data-stu-id="0da3b-112">Override</span></span>](override.md)   | <span data-ttu-id="0da3b-113">Não</span><span class="sxs-lookup"><span data-stu-id="0da3b-113">No</span></span> | <span data-ttu-id="0da3b-114">Especifica a configuração de URLs de localidades adicionais</span><span class="sxs-lookup"><span data-stu-id="0da3b-114">Specifies the setting for additional locale urls</span></span> |

## <a name="attributes"></a><span data-ttu-id="0da3b-115">Atributos</span><span class="sxs-lookup"><span data-stu-id="0da3b-115">Attributes</span></span>

|<span data-ttu-id="0da3b-116">**Atributo**</span><span class="sxs-lookup"><span data-stu-id="0da3b-116">**Attribute**</span></span>|<span data-ttu-id="0da3b-117">**Tipo**</span><span class="sxs-lookup"><span data-stu-id="0da3b-117">**Type**</span></span>|<span data-ttu-id="0da3b-118">**Obrigatório**</span><span class="sxs-lookup"><span data-stu-id="0da3b-118">**Required**</span></span>|<span data-ttu-id="0da3b-119">**Descrição**</span><span class="sxs-lookup"><span data-stu-id="0da3b-119">**Description**</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="0da3b-120">DefaultValue</span><span class="sxs-lookup"><span data-stu-id="0da3b-120">DefaultValue</span></span>|<span data-ttu-id="0da3b-121">URL</span><span class="sxs-lookup"><span data-stu-id="0da3b-121">URL</span></span>|<span data-ttu-id="0da3b-122">obrigatório</span><span class="sxs-lookup"><span data-stu-id="0da3b-122">required</span></span>|<span data-ttu-id="0da3b-123">Especifica o valor padrão para essa configuração, expresso para a localidade especificada no elemento [DefaultLocale](defaultlocale.md).</span><span class="sxs-lookup"><span data-stu-id="0da3b-123">Specifies the default value for this setting, expressed for the locale specified in the [DefaultLocale](defaultlocale.md) element.</span></span>|
