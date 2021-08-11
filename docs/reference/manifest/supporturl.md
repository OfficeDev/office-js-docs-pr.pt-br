---
title: Elemento SupportUrl no arquivo de manifesto
description: O elemento SupportUrl especifica a URL de uma página que fornece informações de suporte para o seu complemento.
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: 1d76afeaaceafc9e8786070338d69cea1b73635d20cd5a729d7e3d859b952494
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 08/07/2021
ms.locfileid: "57096348"
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
