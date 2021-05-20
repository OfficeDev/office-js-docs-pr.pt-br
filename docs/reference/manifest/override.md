---
title: Elemento Override no arquivo de manifesto
description: O elemento Substituir permite especificar o valor de uma configuração dependendo de uma condição especificada.
ms.date: 05/14/2021
localization_priority: Normal
ms.openlocfilehash: 131d72883d050038e2df5b7d8bbca033af9e6ee4
ms.sourcegitcommit: 693d364616b42eea66977eef47530adabc51a40f
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/19/2021
ms.locfileid: "52555154"
---
# <a name="override-element"></a>Elemento Override

Fornece uma maneira de substituir o valor de uma configuração manifesto, dependendo de uma condição especificada. Existem três tipos de condições:

- Uma Office local que é diferente do `LocaleToken` padrão, chamado **LocaleTokenOverride**.
- Um padrão de suporte de conjunto de requisitos diferente do `RequirementToken` padrão padrão, chamado **RequirementTokenOverride**.
- A fonte é diferente do padrão `Runtime` , chamado **RuntimeOverride** (atualmente em pré-visualização).

Um `<Override>` elemento que está dentro de um elemento deve ser do tipo `<Runtime>` **RuntimeOverride**.

Não há `overrideType` atributo para o `<Override>` elemento. A diferença é determinada pelo elemento pai e pelo tipo do elemento pai. Um `<Override>` elemento que está dentro de um elemento cujo é , deve ser do tipo `<Token>` `xsi:type` `RequirementToken` **RequirementTokenOverride**. Um `<Override>` elemento dentro de qualquer outro elemento pai, ou dentro de um elemento do `<Override>` `LocaleToken` tipo, deve ser do tipo **LocaleTokenOverride**. Para obter mais informações sobre o uso desse elemento quando for filho de um `<Token>` elemento, consulte [Trabalho com substituições estendidas do manifesto](../../develop/extended-overrides.md).

Cada tipo é descrito em seções separadas mais tarde neste artigo.

## <a name="override-element-for-localetoken"></a>Elemento de substituição para `LocaleToken`

Um `<Override>` elemento expressa um condicional e pode ser lido como um "Se ... então..." declaração. Se o `<Override>` elemento é do tipo **LocaleTokenOverride,** então o `Locale` atributo é a condição, e o atributo é `Value` o consequente. Por exemplo, o seguinte é "Se o Office configuração local for fr-fr, então o nome de exibição é 'Lecteur vidéo'."

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

## <a name="override-element-for-requirementtoken"></a>Elemento de substituição para `RequirementToken`

Um `<Override>` elemento expressa um condicional e pode ser lido como um "Se ... então..." declaração. Se o `<Override>` elemento é do tipo **RequirementTokenOverride**, então o elemento criança `<Requirements>` expressa a condição, e o atributo é o `Value` consequente. Por exemplo, o primeiro `<Override>` a seguir é lido "Se a plataforma atual suporta o FeatureOne versão 1.7, em seguida, use string 'oldAddinVersion' no lugar do token na URL do avô `${token.requirements}` `<ExtendedOverrides>` (em vez da 'atualização' de string padrão)."

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

|Elemento|Conteúdo|Correio|TaskPane|
|:-----|:-----|:-----|:-----|
|[Requisitos](requirements.md)|||x|

### <a name="attributes"></a>Atributos

|Atributo|Tipo|Obrigatório|Descrição|
|:-----|:-----|:-----|:-----|
|Valor|cadeia de caracteres|obrigatório|Valor do token avó quando a condição está satisfeita.|

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

## <a name="override-element-for-runtime-preview"></a>Elemento de substituição para `Runtime` (visualização)

> [!IMPORTANT]
> Este recurso só é suportado para [visualização](../../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) em Outlook na web e em Windows com uma assinatura Microsoft 365. Para obter mais detalhes, consulte [Configurar seu Outlook complemento para ativação baseada em eventos](../../outlook/autolaunch.md).
>
> Como os recursos de visualização estão sujeitos a alterações sem aviso prévio, eles não devem ser usados em complementos de produção.

Um `<Override>` elemento expressa um condicional e pode ser lido como um "Se ... então..." declaração. Se o `<Override>` elemento é do tipo **RuntimeOverride,** então o `type` atributo é a condição, e o atributo é `resid` o consequente. Por exemplo, o seguinte é "Se o tipo é 'javascript', então o `resid` é 'JSRuntime.Url'." Outlook A área de trabalho requer esse elemento para manipuladores [de pontos de extensão LaunchEvent.](../../reference/manifest/extensionpoint.md#launchevent-preview)

```xml
<Runtime resid="WebViewRuntime.Url">
  <Override type="javascript" resid="JSRuntime.Url"/>
</Runtime>
```

**Tipo de suplemento:** Email

### <a name="syntax"></a>Sintaxe

```XML
<Override type="javascript" resid="JSRuntime.Url"/>
```

### <a name="contained-in"></a>Contido em

- [Tempo de execução](runtime.md)

### <a name="attributes"></a>Atributos

|Atributo|Tipo|Obrigatório|Descrição|
|:-----|:-----|:-----|:-----|
|**type**|string|Sim|Especifica o idioma para esta substituição. No momento, `"javascript"` é a única opção suportada.|
|**resid**|string|Sim|Especifica a localização do URL do arquivo JavaScript que deve substituir a localização url do HTML padrão definido no elemento [Runtime](runtime.md) do pai `resid` . O `resid` pode não ter mais de 32 caracteres e deve corresponder a um atributo de um elemento no `id` `Url` `Resources` elemento.|

### <a name="examples"></a>Exemplos

```xml
<!-- Event-based activation happens in a lightweight runtime.-->
<Runtimes>
  <!-- HTML file including reference to or inline JavaScript event handlers.
  This is used by Outlook on the web. -->
  <Runtime resid="WebViewRuntime.Url">
    <!-- JavaScript file containing event handlers. This is used by Outlook Desktop. -->
    <Override type="javascript" resid="JSRuntime.Url"/>
  </Runtime>
</Runtimes>
```

### <a name="see-also"></a>Confira também

- [Tempo de execução](runtime.md)
- [Configure seu Outlook complemento para ativação baseada em eventos](../../outlook/autolaunch.md)
