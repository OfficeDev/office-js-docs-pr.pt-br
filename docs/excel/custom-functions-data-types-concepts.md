---
title: Funções personalizadas e tipos de dados
description: Use os tipos de dados do Excel com suas funções personalizadas e Suplementos do Office.
ms.date: 12/27/2021
ms.topic: conceptual
ms.custom: scenarios:getting-started
ms.localizationpriority: high
ms.openlocfilehash: 361be486ee45cae87b5cd66e2099dc939418a491
ms.sourcegitcommit: 0be4cd0680d638cf96c12263a71af59ff9f51f5a
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 08/24/2022
ms.locfileid: "67422870"
---
# <a name="use-data-types-with-custom-functions-in-excel-preview"></a>Usar tipos de dados com funções personalizadas no Excel (visualização)

[!include[Custom functions and data types availability note](../includes/excel-custom-functions-data-types-note.md)]

Os tipos de dados expandem a API JavaScript do Excel para dar suporte a tipos de dados além dos quatro tipos de valor de célula originais (cadeia de caracteres, número, booliano e erro). Os tipos de dados incluem suporte para imagens da Web, valores de número formatados, valores de entidade e matrizes nos valores da entidade.

Esses tipos de dados ampliam o poder das funções personalizadas, pois as funções personalizadas aceitam tipos de dados como valores de entrada e saída. Você pode gerar tipos de dados por meio de funções personalizadas ou levar os tipos de dados existentes como argumentos de função nos cálculos. Depois que o esquema JSON de um tipo de dados é definido, esse esquema é mantido em todos os cálculos.

Para saber mais sobre como usar tipos de dados com um suplemento do Excel, consulte [Exibição de tipos de dados em suplementos do Excel](excel-data-types-overview.md).

## <a name="how-custom-functions-handle-data-types"></a>Como as funções personalizadas lidam com tipos de dados

As funções personalizadas podem reconhecer tipos de dados e aceitá-los como valores de parâmetro. Uma função personalizada pode criar um novo tipo de dados para um valor retornado. As funções personalizadas usam o mesmo esquema JSON para tipos de dados que a API JavaScript do Excel, e esse esquema JSON é mantido conforme as funções personalizadas calculam e avaliam.

> [!NOTE]
> As funções personalizadas não dão suporte à funcionalidade completa dos objetos de erro aprimorados oferecidos pelos tipos de dados. Uma função personalizada pode aceitar um objeto de erro de tipos de dados, mas não será mantida durante o cálculo. No momento, as funções personalizadas só dão suporte aos erros incluídos no [objeto CustomFunctions.Error](custom-functions-errors.md).

## <a name="enable-data-types-for-custom-functions"></a>Habilitar tipos de dados para funções personalizadas

Para usar esse recurso, você precisa atualizar manualmente os metadados JSON. Para mais testes temporários, você pode personalizar as configurações do Script Lab em vez de atualizar manualmente os metadados JSON. As seções a seguir descrevem essas etapas mais detalhadamente.

### <a name="manually-update-json-metadata"></a>Atualizar manualmente os metadados JSON

Projetos de funções personalizadas incluem um arquivo de metadados JSON. Esse arquivo de metadados JSON difere do esquema JSON usado pelas APIs de tipos de dados. Para usar a integração de tipos de dados com funções personalizadas, o arquivo de metadados JSON de funções personalizadas deve ser atualizado manualmente para incluir a propriedade `allowCustomDataForDataTypeAny`. Defina essa propriedade como `true`.

Para obter uma descrição completa do processo de criação manual de JSON, confira [Criar metadados JSON manualmente para funções personalizadas](custom-functions-json.md). Confira [allowCustomDataForDataTypeAny](custom-functions-json.md#allowcustomdatafordatatypeany-preview) para obter detalhes adicionais sobre essa propriedade.

### <a name="script-lab-option"></a>Opção Script Lab

A integração de funções personalizadas com tipos de dados está disponível para teste com o Script Lab, além da atualização manual de metadados JSON descrita na seção anterior. Para saber mais sobre o Script Lab, consulte [Explorar a API JavaScript do Office usando o Script Lab](../overview/explore-with-script-lab.md). Para testar esse recurso com o Script Lab, atualize as configurações usando as etapas a seguir.

1. Abra o painel de tarefas Script Lab **Código**.
1. No canto inferior direito, selecione o botão **Configurações**.
1. Vá para **Configurações do Usuário** e insira `allowCustomDataForDataTypeAny: true`.

![Captura de tela mostrando as etapas para habilitar tipos de dados para funções personalizadas no Script Lab.](../images/custom-functions-script-lab-data-type.png)

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

## <a name="see-also"></a>Confira também

* [Visão geral dos tipos de dados em suplementos do Excel](excel-data-types-overview.md)
* [Conceitos básicos dos tipos de dados do Excel](excel-data-types-concepts.md)
* [Configurar seu Suplemento do Office para usar um runtime compartilhado](../develop/configure-your-add-in-to-use-a-shared-runtime.md)
