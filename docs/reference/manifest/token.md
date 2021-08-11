---
title: Elemento Token no arquivo de manifesto
description: Especifica um token ou curinga que pode ser usado com modelos de URL no manifesto.
ms.date: 11/06/2020
localization_priority: Normal
ms.openlocfilehash: 867bb5bc801b85b63c7815debfaf59c5cee3a8157dc866ba7082803ee1d7fe2a
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 08/07/2021
ms.locfileid: "57095903"
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

|Elemento|Conteúdo|Email|TaskPane|
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