---
title: Elemento tokens no arquivo de manifesto
description: Especifica tokens ou curingas que podem ser usados com modelos de URL no manifesto.
ms.date: 11/06/2020
localization_priority: Normal
ms.openlocfilehash: a50de7c2c3e8ebeb9425c1677a94bbcc62281d3b
ms.sourcegitcommit: ca66ff7462bfdf4ed7ae04f43d1388c24de63bf9
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 11/11/2020
ms.locfileid: "48996667"
---
# <a name="tokens-element"></a>Elemento tokens

Define tokens que podem ser usados em URLs de modelo.

**Tipo de suplemento:** Painel de tarefas

## <a name="syntax"></a>Sintaxe

```XML
<Tokens></Tokens>
```

## <a name="contained-in"></a>Contido em

[ExtendedOverrides](extendedoverrides.md)

## <a name="must-contain"></a>Deve conter

|Elemento|Conteúdo|Email|TaskPane|
|:-----|:-----|:-----|:-----|
|[Token](token.md)|||x|

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