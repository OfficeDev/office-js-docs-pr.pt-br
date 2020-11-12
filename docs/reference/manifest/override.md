---
title: Elemento Override no arquivo de manifesto
description: O elemento override permite que você especifique o valor de uma configuração, dependendo de uma condição especificada.
ms.date: 11/06/2020
localization_priority: Normal
ms.openlocfilehash: 2c66503f9f95155a096b1b6fb23332eed8422da6
ms.sourcegitcommit: ca66ff7462bfdf4ed7ae04f43d1388c24de63bf9
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 11/11/2020
ms.locfileid: "48996309"
---
# <a name="override-element"></a>Elemento Override

Fornece uma maneira de substituir o valor de uma configuração de manifesto, dependendo de uma condição especificada. Há dois tipos de condições:

- Uma localidade do Office diferente do padrão.
- Um padrão de suporte ao conjunto de requisitos diferente do padrão padrão.

Há dois tipos de `<Override>` elementos, um é para substituições de localidade, chamado **LocaleTokenOverride** , e o outro para o conjunto de requisitos substitui, chamado **RequirementTokenOverride**. Mas não há nenhum `type` parâmetro para o `<Override>` elemento. A diferença é determinada pelo elemento pai e o tipo do elemento pai. Um `<Override>` elemento que está dentro de um `<Token>` elemento `xsi:type` , cujo é `RequirementToken` , deve ser do tipo **RequirementTokenOverride**. Um `<Override>` elemento dentro de qualquer outro elemento pai ou dentro `<Override>` de um elemento do tipo `LocaleToken` deve ser do tipo **LocaleTokenOverride**. Cada tipo é descrito em seções separadas abaixo.

## <a name="override-element-of-type-localetokenoverride"></a>Elemento override do tipo LocaleTokenOverride

Um `<Override>` elemento expressa um condicional e pode ser lido como um "If... Then... " instrução. Se o `<Override>` elemento for do tipo **LocaleTokenOverride** , o `Locale` atributo será a condição e o `Value` atributo será o consequent. Por exemplo, o seguinte é lido "se a configuração de localidade do Office for fr-fr, o nome para exibição será" Lecteur Vidéo "."

```xml
<DisplayName DefaultValue="Video player">
    <Override Locale="fr-fr" Value="Lecteur vidéo" />
</DisplayName>
```

**Tipo de suplemento:** Conteúdo, Painel de tarefas, Email

### <a name="syntax"></a>Sintaxe

```XML
<Override Locale="string" Value="string"></Override>
```

### <a name="contained-in"></a>Contido em

|Elemento|
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
|[Token](token.md)|

### <a name="attributes"></a>Atributos

|Atributo|Tipo|Obrigatório|Descrição|
|:-----|:-----|:-----|:-----|
|Locale|cadeia de caracteres|obrigatório|Especifica o nome da cultura da localidade para essa substituição no formato de marca do idioma BCP 47, como `"en-US"`.|
|Valor|cadeia de caracteres|obrigatório|Especifica o valor da configuração expressa para a localidade especificada.|

### <a name="examples"></a>Exemplos

```xml
<DisplayName DefaultValue="Video player">
    <Override Locale="fr-fr" Value="Lecteur vidéo" />
</DisplayName>
```

```xml
<bt:Image id="icon1_16x16" DefaultValue="https://www.contoso.com/icon_default.png">
    <bt:Override Locale="ja-jp" Value="https://www.contoso.com/ja-jp16-icon_default.png" />
</bt:Image>
```

```xml
  <ExtendedOverrides Url="http://contoso.com/addinmetadata/${token.locale}/extended-manifest-overrides.json">
    <Tokens>
      <Token Name="locale" DefaultValue="en-us" xsi:type="LocaleToken">
        <Override Locale="es-*" Value="es-es" />
        <Override Locale="es-mx" Value="es-mx" />
        <Override Locale="fr-*" Value="fr-fr" />
        <Override Locale="ja-jp" Value="ja-jp" />
      </Token>
    <Tokens>
  </ExtendedOverrides>
```

### <a name="see-also"></a>Confira também

- [Localização para suplementos do Office](../../develop/localization.md)
- [Atalhos de teclado para o SharePoint](../../design/keyboard-shortcuts.md)

## <a name="override-element-of-type-requirementtokenoverride"></a>Elemento override do tipo RequirementTokenOverride

Um `<Override>` elemento expressa um condicional e pode ser lido como um "If... Then... " instrução. Se o `<Override>` elemento for do tipo **RequirementTokenOverride** , o elemento filho `<Requirements>` expressará a condição e o `Value` atributo será o consequent. Por exemplo, o primeiro `<Override>` no seguinte é lido "se a plataforma atual suportar o FeatureOne versão 1,7, use a cadeia de caracteres ' oldAddinVersion ' no lugar do `${token.requirements}` token na URL do avô `<ExtendedOverrides>` (em vez da cadeia de caracteres padrão" upgrade ")."

```xml
<ExtendedOverrides Url="http://contoso.com/addinmetadata/${token.requirements}/extended-manifest-overrides.json">
    <Tokens>
        <Token Name="requirements" DefaultValue="upgrade" xsi:type="RequirementsToken">
            <Override Value="oldAddinVersion">
                <Requirements>
                    <Sets>
                        <Set Name="FeatureOne" MinVersion="1.7" />
                    </Sets>
                </Requirements>
            </Override>
            <Override Value="currentAddinVersion">
                <Requirements>
                    <Sets>
                        <Set Name="FeatureOne" MinVersion="1.8" />
                    </Sets>
                    <Methods>
                        <Method Name="MethodThree" />
                    </Methods>
                </Requirements>
            </Override>
        </Token>
    </Tokens>
</ExtendedOverrides>
```

**Tipo de suplemento:** Painel de tarefas

### <a name="syntax"></a>Sintaxe

```XML
<Override Value="string" />
```

### <a name="contained-in"></a>Contido em

|Elemento|
|:-----|
|[Token](token.md)|

### <a name="must-contain"></a>Deve conter

|Elemento|Conteúdo|Email|TaskPane|
|:-----|:-----|:-----|:-----|
|[Requisitos](requirements.md)|||x|

### <a name="attributes"></a>Atributos

|Atributo|Tipo|Obrigatório|Descrição|
|:-----|:-----|:-----|:-----|
|Valor|cadeia de caracteres|obrigatório|Valor do token avô quando a condição for atendida.|

### <a name="example"></a>Exemplo

```xml
<ExtendedOverrides Url="http://contoso.com/addinmetadata/${token.requirements}/extended-manifest-overrides.json">
    <Token Name="requirements" DefaultValue="upgrade" xsi:type="RequirementsToken">
        <Override Value="very-old">
            <Requirements>
                <Sets>
                    <Set Name="FeatureOne" MinVersion="1.5" />
                    <Set Name="FeatureTwo" MinVersion="1.1" />
                </Sets>
            </Requirements>
        </Override>
        <Override Value="old">
            <Requirements>
                <Sets>
                    <Set Name="FeatureOne" MinVersion="1.7" />
                    <Set Name="FeatureTwo" MinVersion="1.2" />
                </Sets>
            </Requirements>
        </Override>
        <Override Value="current">
            <Requirements>
                <Sets>
                    <Set Name="FeatureOne" MinVersion="1.8" />
                    <Set Name="FeatureTwo" MinVersion="1.3" />
                </Sets>
                <Methods>
                    <Method Name="MethodThree" />
                </Methods>
            </Requirements>
        </Override>
    </Token>
</ExtendedOverrides>
```

### <a name="see-also"></a>Confira também

- [Versões do Office e conjuntos de requisitos](../../develop/office-versions-and-requirement-sets.md)
- [Definir o elemento Requirements no manifesto](../../develop/specify-office-hosts-and-api-requirements.md#set-the-requirements-element-in-the-manifest)
- [Atalhos de teclado para o SharePoint](../../design/keyboard-shortcuts.md)
