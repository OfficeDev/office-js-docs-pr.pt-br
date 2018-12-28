---
title: Elemento Icon no arquivo de manifesto
description: ''
ms.date: 10/09/2018
ms.openlocfilehash: 476e5720e4959c3c766a7ae6206f2bf12731bfd2
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 12/22/2018
ms.locfileid: "27432267"
---
# <a name="icon-element"></a><span data-ttu-id="e6935-102">Elemento Icon</span><span class="sxs-lookup"><span data-stu-id="e6935-102">Icon element</span></span>

<span data-ttu-id="e6935-103">Define elementos de **Imagem** para controles de [Botão](control.md#button-control) ou de [Menu](control.md#menu-dropdown-button-controls).</span><span class="sxs-lookup"><span data-stu-id="e6935-103">Defines **Image** elements for [Button](control.md#button-control) or [Menu](control.md#menu-dropdown-button-controls) controls.</span></span>

## <a name="attributes"></a><span data-ttu-id="e6935-104">Atributos</span><span class="sxs-lookup"><span data-stu-id="e6935-104">Attributes</span></span>

|  <span data-ttu-id="e6935-105">Atributo</span><span class="sxs-lookup"><span data-stu-id="e6935-105">Attribute</span></span>  |  <span data-ttu-id="e6935-106">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="e6935-106">Required</span></span>  |  <span data-ttu-id="e6935-107">Descrição</span><span class="sxs-lookup"><span data-stu-id="e6935-107">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="e6935-108">**xsi:type**</span><span class="sxs-lookup"><span data-stu-id="e6935-108">**xsi:type**</span></span>  |  <span data-ttu-id="e6935-109">Não</span><span class="sxs-lookup"><span data-stu-id="e6935-109">No</span></span>  | <span data-ttu-id="e6935-p101">O tipo de ícone que está sendo definido. Isso só é aplicável a ícones em fatores forma móveis. Os elementos **Icon** contidos em um elemento [MobileFormFactor](mobileformfactor.md) devem ter esse atributo definido como `bt:MobileIconList`.</span><span class="sxs-lookup"><span data-stu-id="e6935-p101">The type of icon being defined. This is only applicable to icons in mobile form factors. **Icon** elements contained within a [MobileFormFactor](mobileformfactor.md) element must have this attribute set to `bt:MobileIconList`.</span></span> |

## <a name="child-elements"></a><span data-ttu-id="e6935-113">Elementos filho</span><span class="sxs-lookup"><span data-stu-id="e6935-113">Child elements</span></span>

|  <span data-ttu-id="e6935-114">Elemento</span><span class="sxs-lookup"><span data-stu-id="e6935-114">Element</span></span> |  <span data-ttu-id="e6935-115">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="e6935-115">Required</span></span>  |  <span data-ttu-id="e6935-116">Descrição</span><span class="sxs-lookup"><span data-stu-id="e6935-116">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="e6935-117">Image</span><span class="sxs-lookup"><span data-stu-id="e6935-117">Image</span></span>](#image)        | <span data-ttu-id="e6935-118">Sim</span><span class="sxs-lookup"><span data-stu-id="e6935-118">Yes</span></span> |   <span data-ttu-id="e6935-119">resid de uma imagem a usar</span><span class="sxs-lookup"><span data-stu-id="e6935-119">resid of an image to use</span></span>         |

### <a name="image"></a><span data-ttu-id="e6935-120">Image</span><span class="sxs-lookup"><span data-stu-id="e6935-120">Image</span></span>

<span data-ttu-id="e6935-p102">Uma imagem para o botão. O atributo **resid** deve ser definido para o valor do atributo **id** de um elemento **Image** no elemento **Images** no elemento [Resources](resources.md). O atributo **tamanho** indica o tamanho em pixels da imagem. São obrigatórios três tamanhos de imagem (16, 32 e 80 pixels) e há suporte para outros cinco tamanhos (20, 24, 40, 48 e 64 pixels).|</span><span class="sxs-lookup"><span data-stu-id="e6935-p102">An image for the button. The  **resid** attribute must be set to the value of the **id** attribute of an **Image** element in the **Images** element in the [Resources](resources.md) element. The **size** attribute indicates the size in pixels of the image. Three image sizes are required (16, 32, and 80 pixels) while five other sizes are supported (20, 24, 40, 48, and 64 pixels).|</span></span>

```xml
<Icon>
  <bt:Image size="16" resid="blue-icon-16" />
  <bt:Image size="32" resid="blue-icon-32" />
  <bt:Image size="80" resid="blue-icon-80" />
</Icon>
```

## <a name="additional-requirements-for-mobile-form-factors"></a><span data-ttu-id="e6935-125">Requisitos adicionais para fatores forma móveis</span><span class="sxs-lookup"><span data-stu-id="e6935-125">Additional requirements for mobile form factors</span></span>

<span data-ttu-id="e6935-p103">Quando o elemento **Icon** pai é descendente de um elemento [MobileFormFactor](mobileformfactor.md), os tamanhos mínimos necessários são ligeiramente diferentes. O manifesto deve fornecer no mínimo tamanhos de pixel 25, 32 e 48. Cada tamanho fornecido deve aparecer três vezes, com um atributo `scale` definido como `1`, `2` ou `3`.</span><span class="sxs-lookup"><span data-stu-id="e6935-p103">When the parent **Icon** element is a descendant of a [MobileFormFactor](mobileformfactor.md) element, the minimum required sizes are slightly different. The manifest must minimally provide 25, 32, and 48 pixel sizes. Each size provided must appear three times, with a `scale` attribute set to `1`, `2`, or `3`.</span></span>

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