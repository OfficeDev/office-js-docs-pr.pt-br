---
ms.date: 12/18/2019
description: Retornar vários resultados de sua função personalizada em um suplemento do Office Excel.
title: Retornar vários resultados de sua função personalizada
localization_priority: Normal
ms.openlocfilehash: 753755b481ab3db0de711c80ef082aedc82177ae
ms.sourcegitcommit: 682d18c9149b1153f9c38d28e2a90384e6a261dc
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/13/2020
ms.locfileid: "44217834"
---
# <a name="return-multiple-results-from-your-custom-function"></a>Retornar vários resultados de sua função personalizada

Você pode retornar vários resultados de sua função personalizada que serão retornadas às células vizinhas. Esse comportamento é chamado de despejo. Quando sua função personalizada retorna uma matriz de resultados, ela é conhecida como uma fórmula de matriz dinâmica. Para obter mais informações sobre fórmulas de matriz dinâmicas no Excel, consulte [matrizes dinâmicas e comportamento de matriz despejada](https://support.microsoft.com/office/205c6b06-03ba-4151-89a1-87a7eb36e531).

A imagem a seguir mostra como a `SORT` função é despejada para baixo nas células vizinhas. Sua função personalizada também pode retornar vários resultados como este.

![Captura de tela da função "SORT" exibindo vários resultados em várias células.](../images/dynamic-array-spill.png)

Para criar uma função personalizada que seja uma fórmula de matriz dinâmica, ela deve retornar uma matriz bidimensional de valores. Se os resultados forem despejados em células vizinhas que já possuem valores, a fórmula exibirá um `#SPILL!` erro.

O exemplo a seguir mostra como retornar uma matriz dinâmica que derrama.

```javascript
/**
 * Get text values that spill down.
 * @customfunction
 * @returns {string[][]} A dynamic array with multiple results.
 */
function spillDown() {
  return [['first'], ['second'], ['third']];
}
```

O exemplo a seguir mostra como retornar uma matriz dinâmica que despeja à direita. 

```javascript
/**
 * Get text values that spill to the right.
 * @customfunction
 * @returns {string[][]} A dynamic array with multiple results.
 */
function spillRight() {
  return [['first', 'second', 'third']];
}
```

O exemplo a seguir mostra como retornar uma matriz dinâmica que é despejada para baixo e para a direita.

```javascript
/**
 * Get text values that spill both right and down.
 * @customfunction
 * @returns {string[][]} A dynamic array with multiple results.
 */
function spillRectangle() {
  return [
    ['apples', 1, 'pounds'],
    ['oranges', 3, 'pounds'],
    ['pears', 5, 'crates']
  ];
}
```

## <a name="see-also"></a>Confira também

- [Matrizes dinâmicas e comportamento de matriz derramada](https://support.microsoft.com/office/205c6b06-03ba-4151-89a1-87a7eb36e531)
- [Opções para funções personalizadas do Excel](custom-functions-parameter-options.md)