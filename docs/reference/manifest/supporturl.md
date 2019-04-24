---
title: Elemento SupportUrl no arquivo de manifesto
description: ''
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: 18b9b7c4df9def70ab42ae213066188ac04c07a7
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 04/24/2019
ms.locfileid: "32450412"
---
# <a name="supporturl-element"></a><span data-ttu-id="c21f6-102">Elemento SupportUrl</span><span class="sxs-lookup"><span data-stu-id="c21f6-102">SupportUrl element</span></span>

<span data-ttu-id="c21f6-103">Especifica a URL de uma página que fornece informações de suporte para o suplemento.</span><span class="sxs-lookup"><span data-stu-id="c21f6-103">Specifies the URL of a page that provides support information for your add-in.</span></span>

## <a name="syntax"></a><span data-ttu-id="c21f6-104">Sintaxe</span><span class="sxs-lookup"><span data-stu-id="c21f6-104">Syntax</span></span>

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

## <a name="contained-in"></a><span data-ttu-id="c21f6-105">Contido em</span><span class="sxs-lookup"><span data-stu-id="c21f6-105">Contained in</span></span>

[<span data-ttu-id="c21f6-106">OfficeApp</span><span class="sxs-lookup"><span data-stu-id="c21f6-106">OfficeApp</span></span>](officeapp.md)

## <a name="can-contain"></a><span data-ttu-id="c21f6-107">Pode conter</span><span class="sxs-lookup"><span data-stu-id="c21f6-107">Can contain</span></span>

|  <span data-ttu-id="c21f6-108">Elemento</span><span class="sxs-lookup"><span data-stu-id="c21f6-108">Element</span></span> | <span data-ttu-id="c21f6-109">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="c21f6-109">Required</span></span> | <span data-ttu-id="c21f6-110">Descrição</span><span class="sxs-lookup"><span data-stu-id="c21f6-110">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="c21f6-111">Override</span><span class="sxs-lookup"><span data-stu-id="c21f6-111">Override</span></span>](override.md)   | <span data-ttu-id="c21f6-112">Não</span><span class="sxs-lookup"><span data-stu-id="c21f6-112">No</span></span> | <span data-ttu-id="c21f6-113">Especifica a configuração de URLs de localidades adicionais</span><span class="sxs-lookup"><span data-stu-id="c21f6-113">Specifies the setting for additional locale urls</span></span> |

## <a name="attributes"></a><span data-ttu-id="c21f6-114">Atributos</span><span class="sxs-lookup"><span data-stu-id="c21f6-114">Attributes</span></span>

|<span data-ttu-id="c21f6-115">**Atributo**</span><span class="sxs-lookup"><span data-stu-id="c21f6-115">**Attribute**</span></span>|<span data-ttu-id="c21f6-116">**Tipo**</span><span class="sxs-lookup"><span data-stu-id="c21f6-116">**Type**</span></span>|<span data-ttu-id="c21f6-117">**Obrigatório**</span><span class="sxs-lookup"><span data-stu-id="c21f6-117">**Required**</span></span>|<span data-ttu-id="c21f6-118">**Descrição**</span><span class="sxs-lookup"><span data-stu-id="c21f6-118">**Description**</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="c21f6-119">DefaultValue</span><span class="sxs-lookup"><span data-stu-id="c21f6-119">DefaultValue</span></span>|<span data-ttu-id="c21f6-120">URL</span><span class="sxs-lookup"><span data-stu-id="c21f6-120">URL</span></span>|<span data-ttu-id="c21f6-121">obrigatório</span><span class="sxs-lookup"><span data-stu-id="c21f6-121">required</span></span>|<span data-ttu-id="c21f6-122">Especifica o valor padrão para essa configuração, expresso para a localidade especificada no elemento [DefaultLocale](defaultlocale.md).</span><span class="sxs-lookup"><span data-stu-id="c21f6-122">Specifies the default value for this setting, expressed for the locale specified in the [DefaultLocale](defaultlocale.md) element.</span></span>|
