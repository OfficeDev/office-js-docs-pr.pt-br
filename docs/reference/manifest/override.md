---
title: Elemento Override no arquivo de manifesto
description: ''
ms.date: 10/09/2018
ms.openlocfilehash: d1d2400312f12116b1ac5f4010135541e783dcc7
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 12/22/2018
ms.locfileid: "27432862"
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
|[Description](description.md)|
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

- [Localização para suplementos do Office](https://docs.microsoft.com/office/dev/add-ins/develop/localization)
    
