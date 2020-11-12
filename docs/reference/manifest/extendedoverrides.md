---
title: Elemento ExtendedOverrides no arquivo de manifesto
description: Especifica as URLs para uma extensão formatada por JSON do manifesto.
ms.date: 11/06/2020
localization_priority: Normal
ms.openlocfilehash: 76491af34d1caf0ec266826df97a5363e336b85d
ms.sourcegitcommit: ca66ff7462bfdf4ed7ae04f43d1388c24de63bf9
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 11/11/2020
ms.locfileid: "48996670"
---
# <a name="extendedoverrides-element"></a>Elemento ExtendedOverrides

Especifica as URLs completas para arquivos formatados por JSON que estendem o manifesto.

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
|[Sinais](tokens.md)|||x|

## <a name="attributes"></a>Atributos

|Atributo|Descrição|
|:-----|:-----|
|URL (obrigatório)| A URL completa do arquivo JSON de substituições estendidas. Pode ser um modelo de URL que usa tokens definidos pelo elemento [tokens](tokens.md) .|
|ResourcesUrl (opcional) | A URL completa de um arquivo que fornece recursos suplementares, como cadeias de caracteres localizadas, para o arquivo especificado no `Url` atributo. Pode ser um modelo de URL que usa tokens definidos pelo elemento [tokens](tokens.md) .|

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
