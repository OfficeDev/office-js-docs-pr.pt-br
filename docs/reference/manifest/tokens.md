---
title: Elemento Tokens no arquivo de manifesto
description: Especifica tokens ou curingas que podem ser usados com modelos de URL no manifesto.
ms.date: 11/06/2020
localization_priority: Normal
ms.openlocfilehash: 5d42abab46ecc6e7ab465144f061d26da52c0eb3e2623acd8a8a2912ecc13312
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 08/07/2021
ms.locfileid: "57095781"
---
# <a name="tokens-element"></a>Elemento Tokens

Define tokens que podem ser usados em URLs de modelo. Para obter mais informações sobre o uso desse elemento, consulte [Trabalhar com substituições estendidas do manifesto](../../develop/extended-overrides.md).

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