---
title: Funções personalizadas e tipos de dados
description: Use os tipos de dados do Excel com suas funções personalizadas e Suplementos do Office.
ms.date: 10/17/2022
ms.topic: conceptual
ms.custom: scenarios:getting-started
ms.localizationpriority: medium
ms.openlocfilehash: 6ea2287dbf83a5acc45f64c6f5071e504e66bbce
ms.sourcegitcommit: eca6c16d0bb74bed2d35a21723dd98c6b41ef507
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 10/18/2022
ms.locfileid: "68607426"
---
# <a name="use-data-types-with-custom-functions-in-excel"></a>Usar tipos de dados com funções personalizadas no Excel

Os tipos de dados expandem a API JavaScript do Excel para dar suporte a tipos de dados além dos quatro tipos de valor de célula originais (cadeia de caracteres, número, booliano e erro). Os tipos de dados incluem suporte para imagens da Web, valores de número formatados, entidades e matrizes dentro de entidades.

Esses tipos de dados ampliam o poder das funções personalizadas, pois as funções personalizadas aceitam tipos de dados como valores de entrada e saída. Você pode gerar tipos de dados por meio de funções personalizadas ou levar os tipos de dados existentes como argumentos de função nos cálculos. Depois que o esquema JSON de um tipo de dados é definido, esse esquema é mantido em todos os cálculos.

Para saber mais sobre como usar tipos de dados com um suplemento do Excel, consulte [Exibição de tipos de dados em suplementos do Excel](excel-data-types-overview.md).

## <a name="how-custom-functions-handle-data-types"></a>Como as funções personalizadas lidam com tipos de dados

As funções personalizadas podem reconhecer tipos de dados e aceitá-los como valores de parâmetro. Uma função personalizada pode criar um novo tipo de dados para um valor retornado. As funções personalizadas usam o mesmo esquema JSON para tipos de dados que a API JavaScript do Excel, e esse esquema JSON é mantido conforme as funções personalizadas calculam e avaliam.

> [!NOTE]
> As funções personalizadas não dão suporte à funcionalidade completa dos objetos de erro aprimorados oferecidos pelos tipos de dados. Uma função personalizada pode aceitar um objeto de erro de tipos de dados, mas não será mantida durante o cálculo. No momento, as funções personalizadas só dão suporte aos erros incluídos no [objeto CustomFunctions.Error](custom-functions-errors.md).

## <a name="enable-data-types-for-custom-functions"></a>Habilitar tipos de dados para funções personalizadas

Projetos de funções personalizadas incluem um arquivo de metadados JSON. Esse arquivo de metadados JSON difere do esquema JSON usado pelas APIs de tipos de dados. Para usar a integração de tipos de dados com funções personalizadas, o arquivo de metadados JSON de funções personalizadas deve ser atualizado manualmente para incluir a propriedade `allowCustomDataForDataTypeAny`. Defina essa propriedade como `true`.

Para obter uma descrição completa do processo manual de criação de metadados JSON, consulte [Criar manualmente metadados JSON para funções personalizadas](custom-functions-json.md). Consulte [allowCustomDataForDataTypeAny](custom-functions-json.md#allowcustomdatafordatatypeany) para obter detalhes adicionais sobre essa propriedade.

## <a name="output-a-formatted-number-value"></a>Gerar um valor de número formatado

O exemplo de código a seguir mostra como criar um tipo de dados [FormattedNumberCellValue](/javascript/api/excel/excel.formattednumbercellvalue) com uma função personalizada. A função usa um número básico e uma configuração de formato como parâmetros de entrada e retorna um tipo de dados de valor numérico formatado como a saída.

```js
/**
 * Take a number as the input value and return a formatted number value as the output.
 * @customfunction
 * @param {number} value
 * @param {string} format (e.g. "0.00%")
 * @returns A formatted number value.
 */
function createFormattedNumber(value, format) {
    return {
        type: "FormattedNumber",
        basicValue: value,
        numberFormat: format
    }
}
```

## <a name="input-an-entity-value"></a>Inserir um valor de entidade

O exemplo de código a seguir mostra uma função personalizada que usa um tipo de dados [EntityCellValue](/javascript/api/excel/excel.entitycellvalue) como uma entrada. Se o parâmetro `attribute` for definido como `text`, a função retornará a propriedade `text` do valor da entidade. Caso contrário, a função retornará a propriedade `basicValue` do valor da entidade.

```js
/**
 * Accept an entity value data type as a function input.
 * @customfunction
 * @param {any} value
 * @param {string} attribute
 * @returns {any} The text value of the entity.
 */
function getEntityAttribute(value, attribute) {
    if (value.type == "Entity") {
        if (attribute == "text") {
            return value.text;
        } else {
            return value.properties[attribute].basicValue;
        }
    } else {
        return JSON.stringify(value);
    }
}
```

## <a name="next-steps"></a>Próximas etapas

Para experimentar funções personalizadas e tipos de dados, instale o [Script Lab](../overview/explore-with-script-lab.md) no Excel e experimente os tipos de dados: snippet de funções [personalizadas](https://github.com/OfficeDev/office-js-snippets/blob/prod/samples/excel/16-custom-functions/data-types-custom-functions.yaml) em nossa biblioteca **de Exemplos**.

## <a name="see-also"></a>Confira também

* [Visão geral dos tipos de dados em suplementos do Excel](excel-data-types-overview.md)
* [Conceitos básicos dos tipos de dados do Excel](excel-data-types-concepts.md)
* [Configurar seu Suplemento do Office para usar um runtime compartilhado](../develop/configure-your-add-in-to-use-a-shared-runtime.md)
