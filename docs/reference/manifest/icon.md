---
title: Elemento Icon no arquivo de manifesto
description: Define elementos de Imagem para controles de Botão ou de Menu.
ms.date: 10/30/2021
ms.localizationpriority: medium
ms.openlocfilehash: 54ae88e5dceeffa244780764711b263ceabd828d
ms.sourcegitcommit: 23ce57b2702aca19054e31fcb2d2f015b4183ba1
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 11/02/2021
ms.locfileid: "60681723"
---
# <a name="icon-element"></a>Elemento Icon

Define elementos de **Imagem** para controles de [Botão](control.md#button-control) ou de [Menu](control.md#menu-dropdown-button-controls).

## <a name="attributes"></a>Atributos

|  Atributo  |  Obrigatório  |  Descrição  |
|:-----|:-----|:-----|
|  **xsi:type**  |  Não  | O tipo de ícone que está sendo definido. Isso só é aplicável a ícones em fatores forma móveis. Os elementos **Icon** contidos em um elemento [MobileFormFactor](mobileformfactor.md) devem ter esse atributo definido como `bt:MobileIconList`. |

## <a name="child-elements"></a>Elementos filho

|  Elemento |  Obrigatório  |  Descrição  |
|:-----|:-----|:-----|
|  [Imagem](#image)        | Sim |   resid de uma imagem a usar         |

### <a name="image"></a>Image

Uma imagem para o botão. O **atributo resid** não pode ter mais de 32 caracteres e deve ser definido como o valor do atributo **id** de um elemento **Image** no elemento **Images** no [elemento Resources.](resources.md) O atributo **tamanho** indica o tamanho em pixels da imagem. Três tamanhos de imagem são necessários (16, 32 e 80 pixels), enquanto cinco outros tamanhos são suportados (20, 24, 40, 48 e 64 pixels).

```xml
<Icon>
  <bt:Image size="16" resid="blue-icon-16" />
  <bt:Image size="32" resid="blue-icon-32" />
  <bt:Image size="80" resid="blue-icon-80" />
</Icon>
```

> [!IMPORTANT]
> Se essa imagem for o ícone representativo do seu complemento, consulte [Create effective listings in AppSource](/office/dev/store/create-effective-office-store-listings#create-an-icon-for-your-add-in) and within Office for size and other requirements.

## <a name="additional-requirements-for-mobile-form-factors"></a>Requisitos adicionais para fatores forma móveis

Quando o elemento **Icon** pai é descendente de um elemento [MobileFormFactor](mobileformfactor.md), os tamanhos mínimos necessários são ligeiramente diferentes. O manifesto deve fornecer no mínimo tamanhos de pixel 25, 32 e 48. Cada tamanho fornecido deve aparecer três vezes, com um atributo `scale` definido como `1`, `2` ou `3`. Esse atributo especifica a `UIScreen.scale` propriedade para dispositivos iOS. Para obter mais informações, consulte [scale](https://developer.apple.com/documentation/uikit/uiscreen/1617836-scale).

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
