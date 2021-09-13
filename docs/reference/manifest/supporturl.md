---
title: Elemento SupportUrl no arquivo de manifesto
description: O elemento SupportUrl especifica a URL de uma página que fornece informações de suporte para o seu complemento.
ms.date: 10/09/2018
ms.localizationpriority: medium
ms.openlocfilehash: 2ea515aa61ed5bf9e22d6316a76fa4b5e51493f3
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 09/12/2021
ms.locfileid: "59148930"
---
# <a name="supporturl-element"></a>Elemento SupportUrl

Especifica a URL de uma página que fornece informações de suporte para o suplemento.

## <a name="syntax"></a>Sintaxe

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

## <a name="contained-in"></a>Contido em

[OfficeApp](officeapp.md)

## <a name="can-contain"></a>Pode conter

|  Elemento | Obrigatório | Descrição  |
|:-----|:-----|:-----|
|  [Override](override.md)   | Não | Especifica a configuração de URLs de localidades adicionais |

## <a name="attributes"></a>Atributos

|Atributo|Tipo|Obrigatório|Descrição|
|:-----|:-----|:-----|:-----|
|DefaultValue|URL|obrigatório|Especifica o valor padrão para essa configuração, expresso para a localidade especificada no elemento [DefaultLocale](defaultlocale.md).|
