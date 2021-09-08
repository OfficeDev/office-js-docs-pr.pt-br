---
ms.date: 05/11/2020
description: Retornar vários resultados de sua função personalizada em um Office Excel de usuário.
title: Retornar vários resultados de sua função personalizada
localization_priority: Normal
ms.openlocfilehash: b7df6b2c5ca3dca24615a61e11277ac36b42c0df
ms.sourcegitcommit: 42c55a8d8e0447258393979a09f1ddb44c6be884
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 09/08/2021
ms.locfileid: "58939352"
---
# <a name="return-multiple-results-from-your-custom-function"></a>Retornar vários resultados de sua função personalizada

Você pode retornar vários resultados de sua função personalizada que serão retornados para células vizinhas. Esse comportamento é chamado de vazamento. Quando sua função personalizada retorna uma matriz de resultados, ela é conhecida como uma fórmula de matriz dinâmica. Para obter mais informações sobre fórmulas de matriz dinâmicas Excel, consulte [Matrizes dinâmicas e comportamento de matriz descarada](https://support.microsoft.com/office/205c6b06-03ba-4151-89a1-87a7eb36e531).

A imagem a seguir mostra como `SORT` a função é derramada para baixo em células vizinhas. Sua função personalizada também pode retornar vários resultados como este.

![Captura de tela da função 'SORT' exibindo vários resultados para baixo em várias células.](../images/dynamic-array-spill.png)

Para criar uma função personalizada que seja uma fórmula de matriz dinâmica, ela deve retornar uma matriz bidimensional de valores. Se os resultados vazarem em células vizinhas que já têm valores, a fórmula exibirá um `#SPILL!` erro.

O exemplo a seguir mostra como retornar uma matriz dinâmica que se espalha.

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

O exemplo a seguir mostra como retornar uma matriz dinâmica que se espalha à direita. 

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

O exemplo a seguir mostra como retornar uma matriz dinâmica que derrama para baixo e para a direita.

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

- [Matrizes dinâmicas e comportamento de matriz descarrável](https://support.microsoft.com/office/205c6b06-03ba-4151-89a1-87a7eb36e531)
- [Opções para Excel funções personalizadas](custom-functions-parameter-options.md)