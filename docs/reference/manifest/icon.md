---
title: Elemento Icon no arquivo de manifesto
description: ''
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: 45f3dcda8e74430cf70aa765efc6b3aae0e2b448
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 04/24/2019
ms.locfileid: "32450615"
---
# <a name="icon-element"></a><span data-ttu-id="103db-102">Elemento Icon</span><span class="sxs-lookup"><span data-stu-id="103db-102">Icon element</span></span>

<span data-ttu-id="103db-103">Define elementos de **Imagem** para controles de [Botão](control.md#button-control) ou de [Menu](control.md#menu-dropdown-button-controls).</span><span class="sxs-lookup"><span data-stu-id="103db-103">Defines **Image** elements for [Button](control.md#button-control) or [Menu](control.md#menu-dropdown-button-controls) controls.</span></span>

## <a name="attributes"></a><span data-ttu-id="103db-104">Atributos</span><span class="sxs-lookup"><span data-stu-id="103db-104">Attributes</span></span>

|  <span data-ttu-id="103db-105">Atributo</span><span class="sxs-lookup"><span data-stu-id="103db-105">Attribute</span></span>  |  <span data-ttu-id="103db-106">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="103db-106">Required</span></span>  |  <span data-ttu-id="103db-107">Descrição</span><span class="sxs-lookup"><span data-stu-id="103db-107">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="103db-108">**xsi:type**</span><span class="sxs-lookup"><span data-stu-id="103db-108">**xsi:type**</span></span>  |  <span data-ttu-id="103db-109">Não</span><span class="sxs-lookup"><span data-stu-id="103db-109">No</span></span>  | <span data-ttu-id="103db-p101">O tipo de ícone que está sendo definido. Isso só é aplicável a ícones em fatores forma móveis. Os elementos **Icon** contidos em um elemento [MobileFormFactor](mobileformfactor.md) devem ter esse atributo definido como `bt:MobileIconList`.</span><span class="sxs-lookup"><span data-stu-id="103db-p101">The type of icon being defined. This is only applicable to icons in mobile form factors. **Icon** elements contained within a [MobileFormFactor](mobileformfactor.md) element must have this attribute set to `bt:MobileIconList`.</span></span> |

## <a name="child-elements"></a><span data-ttu-id="103db-113">Elementos filho</span><span class="sxs-lookup"><span data-stu-id="103db-113">Child elements</span></span>

|  <span data-ttu-id="103db-114">Elemento</span><span class="sxs-lookup"><span data-stu-id="103db-114">Element</span></span> |  <span data-ttu-id="103db-115">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="103db-115">Required</span></span>  |  <span data-ttu-id="103db-116">Descrição</span><span class="sxs-lookup"><span data-stu-id="103db-116">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="103db-117">Imagem</span><span class="sxs-lookup"><span data-stu-id="103db-117">Image</span></span>](#image)        | <span data-ttu-id="103db-118">Sim</span><span class="sxs-lookup"><span data-stu-id="103db-118">Yes</span></span> |   <span data-ttu-id="103db-119">resid de uma imagem a usar</span><span class="sxs-lookup"><span data-stu-id="103db-119">resid of an image to use</span></span>         |

### <a name="image"></a><span data-ttu-id="103db-120">Image</span><span class="sxs-lookup"><span data-stu-id="103db-120">Image</span></span>

<span data-ttu-id="103db-p102">Uma imagem para o botão. O atributo **resid** deve ser definido para o valor do atributo **id** de um elemento **Image** no elemento **Images** no elemento [Resources](resources.md). O atributo **tamanho** indica o tamanho em pixels da imagem. São obrigatórios três tamanhos de imagem (16, 32 e 80 pixels) e há suporte para outros cinco tamanhos (20, 24, 40, 48 e 64 pixels).|</span><span class="sxs-lookup"><span data-stu-id="103db-p102">An image for the button. The  **resid** attribute must be set to the value of the **id** attribute of an **Image** element in the **Images** element in the [Resources](resources.md) element. The **size** attribute indicates the size in pixels of the image. Three image sizes are required (16, 32, and 80 pixels) while five other sizes are supported (20, 24, 40, 48, and 64 pixels).|</span></span>

```xml
<Icon>
  <bt:Image size="16" resid="blue-icon-16" />
  <bt:Image size="32" resid="blue-icon-32" />
  <bt:Image size="80" resid="blue-icon-80" />
</Icon>
```

## <a name="additional-requirements-for-mobile-form-factors"></a><span data-ttu-id="103db-125">Requisitos adicionais para fatores forma móveis</span><span class="sxs-lookup"><span data-stu-id="103db-125">Additional requirements for mobile form factors</span></span>

<span data-ttu-id="103db-p103">Quando o elemento **Icon** pai é descendente de um elemento [MobileFormFactor](mobileformfactor.md), os tamanhos mínimos necessários são ligeiramente diferentes. O manifesto deve fornecer no mínimo tamanhos de pixel 25, 32 e 48. Cada tamanho fornecido deve aparecer três vezes, com um atributo `scale` definido como `1`, `2` ou `3`.</span><span class="sxs-lookup"><span data-stu-id="103db-p103">When the parent **Icon** element is a descendant of a [MobileFormFactor](mobileformfactor.md) element, the minimum required sizes are slightly different. The manifest must minimally provide 25, 32, and 48 pixel sizes. Each size provided must appear three times, with a `scale` attribute set to `1`, `2`, or `3`.</span></span>

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
