---
title: Elemento Token no arquivo de manifesto
description: Especifica um token ou curinga que pode ser usado com modelos de URL no manifesto.
ms.date: 11/06/2020
localization_priority: Normal
ms.openlocfilehash: 48078f8211a8fd3f0e3f9d7c3f3aabd1d31b0a6d
ms.sourcegitcommit: e7009c565b18c607fe0868db2e26e250ad308dce
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/05/2021
ms.locfileid: "50505364"
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