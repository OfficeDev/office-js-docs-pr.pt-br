---
title: Elemento Override no arquivo de manifesto
description: O elemento override permite que você especifique o valor de uma configuração para uma localidade adicional.
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: aa5d023169389670d15e36f8bee4445529d84711
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 06/08/2020
ms.locfileid: "44611502"
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
