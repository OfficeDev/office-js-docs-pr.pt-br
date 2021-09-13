---
title: Elemento Tokens no arquivo de manifesto
description: Especifica tokens ou curingas que podem ser usados com modelos de URL no manifesto.
ms.date: 11/06/2020
ms.localizationpriority: medium
ms.openlocfilehash: 3e52543bdb53709ea005f63a3a990650905d70cd
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 09/12/2021
ms.locfileid: "59152077"
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

|Elemento|Conteúdo|Correio|TaskPane|
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