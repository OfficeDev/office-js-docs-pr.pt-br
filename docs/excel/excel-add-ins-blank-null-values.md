---
title: Valores em branco e nulo nos suplementos do Excel
description: Saiba como trabalhar com valores nulos em branco Excel métodos e propriedades do modelo de objeto.
ms.date: 09/03/2020
localization_priority: Normal
ms.openlocfilehash: 3f38569f7342bb88c52ce424db426bfa7939be5e
ms.sourcegitcommit: 42c55a8d8e0447258393979a09f1ddb44c6be884
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 09/08/2021
ms.locfileid: "58937680"
---
# <a name="blank-and-null-values-in-excel-add-ins"></a>Valores em branco e nulo nos suplementos do Excel

`null` e as cadeias de caracteres esvaziadas têm implicações especiais nas APIs JavaScript do Excel. Elas são usadas para representar células vazias, sem formatação ou valores padrão. Essa seção detalha o uso da `null` e de uma cadeia de caracteres vazia ao obter e definir as propriedades.

## <a name="null-input-in-2-d-array"></a>entrada nula em uma matriz 2D

No Excel, um intervalo é representado por uma matriz 2D, onde a primeira dimensão é linhas e a segunda dimensão é colunas. Para definir valores, o formato do número ou a fórmula apenas para células específicas em um intervalo, especifique os valores, o formato do número ou a fórmula para essas células na matriz 2D, bem como `null` para todas as outras células na matriz 2D.

Por exemplo, para atualizar o formato do número apenas para uma célula em um intervalo e manter o formato de número existente para todas as outras células no intervalo, especifique o novo formato de número para a célula a ser atualizada e `null` para todas as outras células. O trecho de código a seguir define um novo formato de número para a quarta célula no intervalo e não altera o formato de número para as primeiras três células no intervalo.

```js
range.values = [['Eurasia', '29.96', '0.25', '15-Feb' ]];
range.numberFormat = [[null, null, null, 'm/d/yyyy;@']];
```

## <a name="null-input-for-a-property"></a>entrada nula para uma propriedade

`null` não é uma entrada válida para uma propriedade única. Por exemplo, o trecho de código a seguir não é válido, pois a propriedade `values` do intervalo não pode ser definida como `null`.

```js
range.values = null; // This is not a valid snippet. 
```

Da mesma forma, o seguinte snippet de código não é válido, pois `null` não é um valor válido para a propriedade `color`.

```js
range.format.fill.color =  null;  // This is not a valid snippet. 
```

## <a name="null-property-values-in-the-response"></a>Valores da propriedade nula na resposta

A formatação de propriedades como `size` e `color` conterá valores `null` na resposta quando valores diferentes existirem no intervalo especificado. Por exemplo, se você recuperar um intervalo e carregar sua propriedade `format.font.color`:

* Se todas as células no intervalo tiverem a mesma cor de fonte, `range.format.font.color` especificará essa cor.
* Se houver várias cores de fonte dentro do intervalo, `range.format.font.color` será `null`.

## <a name="blank-input-for-a-property"></a>Entrada em branco para uma propriedade

Quando você especificar um valor em branco para uma propriedade (isto é, duas aspas sem espaço entre elas `''`), ele será interpretado como uma instrução para limpar ou redefinir a propriedade. Por exemplo:

* Se você especificar um valor em branco para a propriedade `values` de um intervalo, o conteúdo do intervalo será apagado.
* Se você especificar um valor em branco para a propriedade `numberFormat`, o formato de número será redefinido para `General`.
* Se você especificar um valor em branco para a propriedade `formula` e a propriedade `formulaLocale`, os valores de fórmula serão apagados.

## <a name="blank-property-values-in-the-response"></a>Valores da propriedade em branco na resposta

Para operações de leitura, um valor de propriedade em branco na resposta (isto é, duas aspas sem espaço entre elas `''`) indica que a célula não contém dados nem valor. No primeiro exemplo abaixo, a primeira e a última célula no intervalo não contêm dados. No segundo exemplo, as primeiras duas células no intervalo não contêm uma fórmula.

```js
range.values = [['', 'some', 'data', 'in', 'other', 'cells', '']];
```

```js
range.formula = [['', '', '=Rand()']];
```
