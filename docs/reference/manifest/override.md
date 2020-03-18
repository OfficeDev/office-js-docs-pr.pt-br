---
title: Elemento Override no arquivo de manifesto
description: O elemento override permite que você especifique o valor de uma configuração para uma localidade adicional.
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: 39e706dc981d405fcfcc508626578f34931efbcb
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/17/2020
ms.locfileid: "42718024"
---
# <a name="override-element"></a>Elemento Override

Fornece uma maneira de especificar o valor de uma configuração para uma localidade adicional.

**Tipo de suplemento:** Conteúdo, Painel de tarefas, Email

## <a name="syntax"></a>Sintaxe

```XML
<Override Locale="string" Value="string" />
```

## <a name="contained-in"></a>Contido em

|**Elemento**|
|:-----|
|[CitationText](citationtext.md)|
|[Descrição](description.md)|
|[DictionaryName](dictionaryname.md)|
|[DictionaryHomePage](dictionaryhomepage.md)|
|[DisplayName](displayname.md)|
|[HighResolutionIconUrl](highresolutioniconurl.md)|
|[IconUrl](iconurl.md)|
|[QueryUri](queryuri.md)|
|[SourceLocation](sourcelocation.md)|
|[SupportUrl](supporturl.md)|

## <a name="attributes"></a>Atributos

|**Atributo**|**Tipo**|**Obrigatório**|**Descrição**|
|:-----|:-----|:-----|:-----|
|Locale|cadeia de caracteres|obrigatório|Especifica o nome da cultura da localidade para essa substituição no formato de marca do idioma BCP 47, como `"en-US"`.|
|Valor|cadeia de caracteres|obrigatório|Especifica o valor da configuração expressa para a localidade especificada.|

## <a name="see-also"></a>Confira também

- [Localização para suplementos do Office](../../develop/localization.md)
