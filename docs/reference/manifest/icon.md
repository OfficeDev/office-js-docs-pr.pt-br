---
title: Elemento Icon no arquivo de manifesto
description: Define elementos de Imagem para controles de Botão ou de Menu.
ms.date: 03/30/2021
localization_priority: Normal
ms.openlocfilehash: 1adfbcd154091fcae49966f0c1f7d0b9cc968ed3
ms.sourcegitcommit: 0bff0411d8cfefd4bb00c189643358e6fb1df95e
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 04/07/2021
ms.locfileid: "51604621"
---
# <a name="icon-element"></a><span data-ttu-id="1a7b3-103">Elemento Icon</span><span class="sxs-lookup"><span data-stu-id="1a7b3-103">Icon element</span></span>

<span data-ttu-id="1a7b3-104">Define elementos de **Imagem** para controles de [Botão](control.md#button-control) ou de [Menu](control.md#menu-dropdown-button-controls).</span><span class="sxs-lookup"><span data-stu-id="1a7b3-104">Defines **Image** elements for [Button](control.md#button-control) or [Menu](control.md#menu-dropdown-button-controls) controls.</span></span>

## <a name="attributes"></a><span data-ttu-id="1a7b3-105">Atributos</span><span class="sxs-lookup"><span data-stu-id="1a7b3-105">Attributes</span></span>

|  <span data-ttu-id="1a7b3-106">Atributo</span><span class="sxs-lookup"><span data-stu-id="1a7b3-106">Attribute</span></span>  |  <span data-ttu-id="1a7b3-107">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="1a7b3-107">Required</span></span>  |  <span data-ttu-id="1a7b3-108">Descrição</span><span class="sxs-lookup"><span data-stu-id="1a7b3-108">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="1a7b3-109">**xsi:type**</span><span class="sxs-lookup"><span data-stu-id="1a7b3-109">**xsi:type**</span></span>  |  <span data-ttu-id="1a7b3-110">Não</span><span class="sxs-lookup"><span data-stu-id="1a7b3-110">No</span></span>  | <span data-ttu-id="1a7b3-p101">O tipo de ícone que está sendo definido. Isso só é aplicável a ícones em fatores forma móveis. Os elementos **Icon** contidos em um elemento [MobileFormFactor](mobileformfactor.md) devem ter esse atributo definido como `bt:MobileIconList`.</span><span class="sxs-lookup"><span data-stu-id="1a7b3-p101">The type of icon being defined. This is only applicable to icons in mobile form factors. **Icon** elements contained within a [MobileFormFactor](mobileformfactor.md) element must have this attribute set to `bt:MobileIconList`.</span></span> |

## <a name="child-elements"></a><span data-ttu-id="1a7b3-114">Elementos filho</span><span class="sxs-lookup"><span data-stu-id="1a7b3-114">Child elements</span></span>

|  <span data-ttu-id="1a7b3-115">Elemento</span><span class="sxs-lookup"><span data-stu-id="1a7b3-115">Element</span></span> |  <span data-ttu-id="1a7b3-116">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="1a7b3-116">Required</span></span>  |  <span data-ttu-id="1a7b3-117">Descrição</span><span class="sxs-lookup"><span data-stu-id="1a7b3-117">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="1a7b3-118">Imagem</span><span class="sxs-lookup"><span data-stu-id="1a7b3-118">Image</span></span>](#image)        | <span data-ttu-id="1a7b3-119">Sim</span><span class="sxs-lookup"><span data-stu-id="1a7b3-119">Yes</span></span> |   <span data-ttu-id="1a7b3-120">resid de uma imagem a usar</span><span class="sxs-lookup"><span data-stu-id="1a7b3-120">resid of an image to use</span></span>         |

### <a name="image"></a><span data-ttu-id="1a7b3-121">Image</span><span class="sxs-lookup"><span data-stu-id="1a7b3-121">Image</span></span>

<span data-ttu-id="1a7b3-122">Uma imagem para o botão.</span><span class="sxs-lookup"><span data-stu-id="1a7b3-122">An image for the button.</span></span> <span data-ttu-id="1a7b3-123">O **atributo resid** não pode ter mais de 32 caracteres e deve ser definido como o valor do atributo **id** de um elemento **Image** no elemento **Images** no [elemento Resources.](resources.md)</span><span class="sxs-lookup"><span data-stu-id="1a7b3-123">The **resid** attribute can be no more than 32 characters and must be set to the value of the **id** attribute of an **Image** element in the **Images** element in the [Resources](resources.md) element.</span></span> <span data-ttu-id="1a7b3-124">O atributo **tamanho** indica o tamanho em pixels da imagem.</span><span class="sxs-lookup"><span data-stu-id="1a7b3-124">The **size** attribute indicates the size in pixels of the image.</span></span> <span data-ttu-id="1a7b3-125">Três tamanhos de imagem são necessários (16, 32 e 80 pixels), enquanto cinco outros tamanhos são suportados (20, 24, 40, 48 e 64 pixels).</span><span class="sxs-lookup"><span data-stu-id="1a7b3-125">Three image sizes are required (16, 32, and 80 pixels) while five other sizes are supported (20, 24, 40, 48, and 64 pixels).</span></span>

```xml
<Icon>
  <bt:Image size="16" resid="blue-icon-16" />
  <bt:Image size="32" resid="blue-icon-32" />
  <bt:Image size="80" resid="blue-icon-80" />
</Icon>
```

> [!IMPORTANT]
> <span data-ttu-id="1a7b3-126">Se essa imagem for o ícone representativo do seu complemento, consulte [Create effective listings in AppSource](/office/dev/store/create-effective-office-store-listings#create-an-icon-for-your-add-in) and within Office for size and other requirements.</span><span class="sxs-lookup"><span data-stu-id="1a7b3-126">If this image is your add-in's representative icon, see [Create effective listings in AppSource and within Office](/office/dev/store/create-effective-office-store-listings#create-an-icon-for-your-add-in) for size and other requirements.</span></span>

## <a name="additional-requirements-for-mobile-form-factors"></a><span data-ttu-id="1a7b3-127">Requisitos adicionais para fatores forma móveis</span><span class="sxs-lookup"><span data-stu-id="1a7b3-127">Additional requirements for mobile form factors</span></span>

<span data-ttu-id="1a7b3-p103">Quando o elemento **Icon** pai é descendente de um elemento [MobileFormFactor](mobileformfactor.md), os tamanhos mínimos necessários são ligeiramente diferentes. O manifesto deve fornecer no mínimo tamanhos de pixel 25, 32 e 48. Cada tamanho fornecido deve aparecer três vezes, com um atributo `scale` definido como `1`, `2` ou `3`.</span><span class="sxs-lookup"><span data-stu-id="1a7b3-p103">When the parent **Icon** element is a descendant of a [MobileFormFactor](mobileformfactor.md) element, the minimum required sizes are slightly different. The manifest must minimally provide 25, 32, and 48 pixel sizes. Each size provided must appear three times, with a `scale` attribute set to `1`, `2`, or `3`.</span></span>

```xml
<Icon xsi:type="bt:MobileIconList">
  <bt:Image resid="blue-icon-16-1" size="25" scale="1" />
  <bt:Image resid="blue-icon-16-2" size="25" scale="2" />
  <bt:Image resid="blue-icon-16-3" size="25" scale="3" />
  <bt:Image resid="blue-icon-32-1" size="32" scale="1" />
  <bt:Image resid="blue-icon-32-2" size="32" scale="2" />
  <bt:Image resid="blue-icon-32-3" size="32" scale="3" />
  <bt:Image resid="blue-icon-80-1" size="48" scale="1" />
  <bt:Image resid="blue-icon-80-2" size="48" scale="2" />
  <bt:Image resid="blue-icon-80-3" size="48" scale="3" />
</Icon>
```
