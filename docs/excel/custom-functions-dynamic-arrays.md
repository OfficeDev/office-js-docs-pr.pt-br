---
ms.date: 05/11/2020
description: Retorne vários resultados de sua função personalizada em um Office Excel de usuário.
title: Retornar vários resultados de sua função personalizada
ms.localizationpriority: medium
ms.openlocfilehash: 9c619b379bc39598bb325180d32ddcbced0ff664
ms.sourcegitcommit: 968d637defe816449a797aefd930872229214898
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/23/2022
ms.locfileid: "63744354"
---
# <a name="return-multiple-results-from-your-custom-function"></a>Retornar vários resultados de sua função personalizada

Você pode retornar vários resultados de sua função personalizada que serão retornados para células vizinhas. Esse comportamento é chamado de vazamento. Quando sua função personalizada retorna uma matriz de resultados, ela é conhecida como uma fórmula de matriz dinâmica. Para obter mais informações sobre fórmulas de matriz dinâmicas Excel, consulte [Matrizes dinâmicas e comportamento de matrizes descarada](https://support.microsoft.com/office/205c6b06-03ba-4151-89a1-87a7eb36e531).

A imagem a seguir mostra como a `SORT` função é derramada para baixo em células vizinhas. Sua função personalizada também pode retornar vários resultados como este.

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
