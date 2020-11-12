---
title: Elemento token no arquivo de manifesto
description: Especifica um token ou curinga que pode ser usado com modelos de URL no manifesto.
ms.date: 11/06/2020
localization_priority: Normal
ms.openlocfilehash: 5e26af44c566ab09ac81c8194e1ae7d85aaac327
ms.sourcegitcommit: ca66ff7462bfdf4ed7ae04f43d1388c24de63bf9
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 11/11/2020
ms.locfileid: "48996669"
---
# <a name="token-element"></a>Elemento token

Define um token de URL individual.

**Tipo de suplemento:** Painel de tarefas

## <a name="syntax"></a>Sintaxe

```XML
<Token Name="string" DefaultValue="string" xsi:type=["LocaleToken" | "RequirementsToken"] ></Token>
```

## <a name="contained-in"></a>Contido em

[Sinais](tokens.md)

## <a name="can-contain"></a>Pode conter

|Elemento|Conteúdo|Email|TaskPane|
|:-----|:-----|:-----|:-----|
|[Override](override.md)|||x|

## <a name="attributes"></a>Atributos

|Atributo|Descrição|
|:-----|:-----|
|DefaultValue|Valor padrão para esse token se nenhuma condição em qualquer `<Override>` elemento filho corresponder.|
|Nome|Nome do token. Esse nome é definido pelo usuário. O tipo do token é determinado pelo atributo Type.|
|xsi:type|Define o tipo de token. Este atributo deve ser definido como um de:  `"RequirementsToken"` ou  `"LocaleToken"` .|

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