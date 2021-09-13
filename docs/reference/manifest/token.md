---
title: Elemento Token no arquivo de manifesto
description: Especifica um token ou curinga que pode ser usado com modelos de URL no manifesto.
ms.date: 11/06/2020
ms.localizationpriority: medium
ms.openlocfilehash: 69f626f5f6f57dd155756812bcd56267a1da3ffa
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 09/12/2021
ms.locfileid: "59148918"
---
# <a name="token-element"></a>Elemento Token

Define um token de URL individual. Para obter mais informações sobre o uso desse elemento, consulte [Trabalhar com substituições estendidas do manifesto](../../develop/extended-overrides.md).

**Tipo de suplemento:** Painel de tarefas

## <a name="syntax"></a>Sintaxe

```XML
<Token Name="string" DefaultValue="string" xsi:type=["LocaleToken" | "RequirementsToken"] ></Token>
```

## <a name="contained-in"></a>Contido em

[Tokens](tokens.md)

## <a name="can-contain"></a>Pode conter

|Elemento|Conteúdo|Correio|TaskPane|
|:-----|:-----|:-----|:-----|
|[Override](override.md)|||x|

## <a name="attributes"></a>Atributos

|Atributo|Descrição|
|:-----|:-----|
|DefaultValue|Valor padrão para esse token se nenhuma condição em qualquer `<Override>` elemento filho corresponde.|
|Nome|Nome do token. Esse nome é definido pelo usuário. O tipo do token é determinado pelo atributo type.|
|xsi:type|Define o tipo de Token. Esse atributo deve ser definido como um dos:  `"RequirementsToken"` , ou  `"LocaleToken"` .|

## <a name="example"></a>Exemplo

```XML
<OfficeApp ...>
  <!-- other elements omitted -->
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
</OfficeApp>
```