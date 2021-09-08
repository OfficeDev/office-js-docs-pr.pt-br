---
title: Elemento ExtendedOverrides no arquivo de manifesto
description: Especifica as URLs para uma extensão formatada por JSON do manifesto.
ms.date: 02/23/2021
localization_priority: Normal
ms.openlocfilehash: f433c9c5604f3fae35580ba20780ea6fe91401c7
ms.sourcegitcommit: 42c55a8d8e0447258393979a09f1ddb44c6be884
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 09/08/2021
ms.locfileid: "58936251"
---
# <a name="extendedoverrides-element"></a>Elemento ExtendedOverrides

Especifica as URLs completas para arquivos formatados com JSON que estendem o manifesto. Para obter informações detalhadas sobre o uso desse elemento e seus elementos [descendentes,](../../develop/extended-overrides.md)consulte Trabalhar com substituições estendidas do manifesto .

**Tipo de suplemento:** Painel de tarefas

## <a name="syntax"></a>Sintaxe

```XML
<ExtendedOverrides Url="string" [ResourcesUrl="string"] ></ExtendedOverrides>
```

## <a name="contained-in"></a>Contido em

[OfficeApp](officeapp.md)

## <a name="can-contain"></a>Pode conter

|Elemento|Conteúdo|Email|TaskPane|
|:-----|:-----|:-----|:-----|
|[Tokens](tokens.md)|||x|

## <a name="attributes"></a>Atributos

|Atributo|Descrição|
|:-----|:-----|
|URL (obrigatório)| A URL completa do arquivo JSON substitui estendido. No futuro, esse valor pode ser um modelo de URL que usa tokens definidos pelo [elemento Tokens.](tokens.md) Consulte [Exemplos](#examples).|
|ResourcesUrl (opcional) | A URL completa de um arquivo que fornece recursos suplementares, como cadeias de caracteres localizadas, para o arquivo especificado no `Url` atributo. Pode ser um modelo de URL que usa tokens definidos pelo [elemento Tokens.](tokens.md)|

## <a name="examples"></a>Exemplos

```XML
<OfficeApp ...>
  <!-- other elements omitted -->
  <ExtendedOverrides Url="http://contoso.com/addinmetadata/extended-manifest-overrides.json"
                     ResourceUrl="https://contoso.com/addin/my-resources.json">
  </ExtendedOverrides>
</OfficeApp>
```

No futuro, esse valor pode ser um modelo de URL que usa tokens definidos pelo [elemento Tokens.](tokens.md) Apresentamos um exemplo a seguir.

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
