---
title: Conceitos principais de funções e tipos de dados personalizados
description: Saiba os principais conceitos para usar Excel de dados com suas funções personalizadas.
ms.date: 11/03/2021
ms.topic: conceptual
ms.custom: scenarios:getting-started
ms.localizationpriority: medium
ms.openlocfilehash: 3b7e735f78ca7b6dcdffa3bd5e8ba9c9d3093766
ms.sourcegitcommit: ad5d7ab21f64012543fb2bd9226d90330d25468b
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 11/04/2021
ms.locfileid: "60749403"
---
# <a name="custom-functions-and-data-types-core-concepts-preview"></a>Conceitos principais de funções e tipos de dados personalizados (visualização)

[!include[Custom functions and data types availability note](../includes/excel-custom-functions-data-types-note.md)]

Os tipos de dados aprimoram Excel API JavaScript expandindo o suporte para tipos de dados além dos quatro originais (cadeia de caracteres, número, booleano e erro). Os tipos de dados incluem suporte para valores de número formatados, imagens da Web, valores de entidade e matrizes dentro dos valores da entidade. Funções personalizadas aceitam tipos de dados como valores de entrada e saída, expandindo o poder de cálculo de funções personalizadas.

Para saber mais sobre como usar tipos de dados com um Excel de Excel, consulte [Excel conceitos principais](excel-data-types-concepts.md)de tipos de dados.

## <a name="how-custom-functions-handle-data-types"></a>Como as funções personalizadas lidam com tipos de dados

Funções personalizadas podem reconhecer tipos de dados e aceitá-los como valores de parâmetro. Uma função personalizada pode criar um novo tipo de dados para um valor de retorno. As funções personalizadas usam o mesmo esquema JSON para tipos de dados que Excel API JavaScript do Excel, e esse esquema JSON é mantido conforme as funções personalizadas calculam e avaliam.

> [!NOTE]
> Funções personalizadas não suportam a funcionalidade completa dos objetos de erro aprimorados oferecidos por tipos de dados. Uma função personalizada pode aceitar um objeto de erro de tipos de dados, mas não será mantida durante o cálculo. No momento, as funções personalizadas só suportam os erros incluídos no [objeto CustomFunctions.Error.](custom-functions-errors.md)

## <a name="enable-data-types-for-custom-functions"></a>Habilitar tipos de dados para funções personalizadas

Para usar esse recurso, você precisa atualizar manualmente seus metadados JSON. Para testes mais temporários, você pode personalizar suas configurações Script Lab em vez de atualizar manualmente os metadados JSON. As seções a seguir detalham essas etapas com mais detalhes.

### <a name="manually-update-json-metadata"></a>Atualizar manualmente metadados JSON

Os projetos de funções personalizadas incluem um arquivo de metadados JSON. Esse arquivo de metadados JSON difere do esquema JSON usado pelas APIs de tipos de dados. Para usar a integração de tipos de dados com funções personalizadas, o arquivo de metadados JSON de funções personalizadas deve ser atualizado manualmente para incluir a propriedade `allowCustomDataForDataTypeAny` . De definir essa propriedade como `true` .

Para uma descrição completa do processo de criação JSON manual, consulte [Manualmente criar metadados JSON para funções personalizadas.](custom-functions-json.md) Consulte [allowCustomDataForDataTypeAny](custom-functions-json.md#allowcustomdatafordatatypeany-preview) para obter detalhes adicionais sobre essa propriedade.

### <a name="script-lab-option"></a>Script Lab opção

A integração de funções personalizadas com tipos de dados está disponível para testes com Script Lab, além da atualização manual de metadados JSON descrita na seção anterior. Para saber mais sobre Script Lab, consulte [Explore Office API JavaScript usando Script Lab](../overview/explore-with-script-lab.md). Para testar esse recurso com Script Lab, atualize as configurações usando as etapas a seguir.

1. Abra o painel de tarefas **Script Lab** Código.
1. No canto inferior direito, selecione o **botão Configurações.**
1. Vá até a **guia Usuário Configurações** insira `allowCustomDataForDataTypeAny: true` .

![Captura de tela mostrando as etapas para habilitar tipos de dados para funções personalizadas Script Lab.](../images/custom-functions-script-lab-data-type.png)

## <a name="output-a-formatted-number-value"></a>Saída de um valor de número formatado

O exemplo de código a seguir mostra como criar um tipo de dados [FormattedNumberCellValue](/javascript/api/excel/excel.formattednumbercellvalue) com uma função personalizada. A função tem um número básico e uma configuração de formato como parâmetros de entrada e retorna um tipo de dados de valor de número formatado como a saída.

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

O exemplo de código a seguir mostra uma função personalizada que leva um tipo de dados [EntityCellValue](/javascript/api/excel/excel.entitycellvalue) como uma entrada. Se o `attribute` parâmetro for definido como , a função `text` retornará a propriedade do valor da `text` entidade. Caso contrário, a função `basicValue` retornará a propriedade do valor da entidade.

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

* [Visão geral de tipos de dados e funções personalizadas](custom-functions-data-types-overview.md)
* [Visão geral dos tipos de dados em suplementos do Excel](excel-data-types-overview.md)
* [Conceitos básicos dos tipos de dados do Excel](excel-data-types-concepts.md)
* [Configure seu Suplemento do Office para usar um tempo de execução de JavaScript compartilhado](../develop/configure-your-add-in-to-use-a-shared-runtime.md)
