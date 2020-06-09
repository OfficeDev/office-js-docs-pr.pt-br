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

|**Atributo**|**Tipo**|**Obrigatório**|**Descrição**|
|:-----|:-----|:-----|:-----|
|DefaultValue|URL|obrigatório|Especifica o valor padrão para essa configuração, expresso para a localidade especificada no elemento [DefaultLocale](defaultlocale.md).|
