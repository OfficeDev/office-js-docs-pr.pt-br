---
ms.date: 10/17/2018
description: Defina metadados para funções personalizadas no Excel.
title: Metadados para funções personalizadas no Excel
ms.openlocfilehash: cff1cbc22f39c99597d4abe7005d7b8bbce6e185
ms.sourcegitcommit: a6d6348075c1abed76d2146ddfc099b0151fe403
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 10/19/2018
ms.locfileid: "25640005"
---
# <a name="custom-functions-metadata-preview"></a>Metadados de funções personalizadas (versão prévia)

Quando você define [funções personalizadas](custom-functions-overview.md) no suplemento do Excel, o projeto de suplemento deve incluir um arquivo de metadados JSON que forneça as informações necessárias para o Excel registrar as funções personalizadas e disponibilizá-las aos usuários finais. Este artigo descreve o formato do arquivo de metadados JSON.

Para obter informações sobre os outros arquivos que você deve incluir no projeto de suplemento para habilitar funções personalizadas, confira [Criar funções personalizadas no Excel](custom-functions-overview.md).

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

## <a name="example-metadata"></a>Exemplo de metadados

O exemplo a seguir mostra o conteúdo de um arquivo de metadados JSON para um suplemento que define funções personalizadas. As seções a seguir neste exemplo fornecem informações detalhadas sobre as propriedades individuais nesse exemplo JSON.

```json
{
  "functions": [
    {
      "id": "ADD",
      "name": "ADD",
      "description": "Add two numbers",
      "helpUrl": "http://www.contoso.com/help",
      "result": {
        "type": "number",
        "dimensionality": "scalar"
      },
      "parameters": [
        {
          "name": "first",
          "description": "first number to add",
          "type": "number",
          "dimensionality": "scalar"
        },
        {
          "name": "second",
          "description": "second number to add",
          "type": "number",
          "dimensionality": "scalar"
        }
      ]
    },
    {
      "id": "GETDAY",
      "name": "GETDAY",
      "description": "Get the day of the week",
      "helpUrl": "http://www.contoso.com/help",
      "result": {
        "type": "string"
      },
      "parameters": []
    },
    {
      "id": "INCREMENTVALUE",
      "name": "INCREMENTVALUE", 
      "description":  "Count up from zero",
      "helpUrl": "http://www.contoso.com/help",
      "result": {
        "type": "number",
        "dimensionality": "scalar"
      },
      "parameters": [
        {
          "name": "increment",
          "description": "the number to be added each time",
          "type": "number",
          "dimensionality": "scalar"
        }
      ],
      "options": {
        "stream": true,
        "cancelable": true
      }
    },
    {
      "id": "SECONDHIGHEST",
      "name": "SECONDHIGHEST", 
      "description":  "Get the second highest number from a range",
      "helpUrl": "http://www.contoso.com/help",
      "result": {
        "type": "number",
        "dimensionality": "scalar"
      },
      "parameters": [
        {
          "name": "range",
          "description": "the input range",
          "type": "number",
          "dimensionality": "matrix"
        }
      ]
    }
  ]
}
```

> [!NOTE]
> Um exemplo completo do arquivo JSON está disponível no [repositório GitHub OfficeDev/Excel-Custom-Functions](https://github.com/OfficeDev/Excel-Custom-Functions/blob/master/config/customfunctions.json).

## <a name="functions"></a>functions 

A `functions` propriedade é uma matriz de objetos de função personalizada. A tabela a seguir lista as propriedades de cada objeto.

|  Propriedade  |  Tipo de dados  |  Obrigatório  |  Descrição  |
|:-----|:-----|:-----|:-----|
|  `description`  |  string  |  Não  |  A descrição da função que os usuários finais veem no Excel. Por exemplo, **Converte um valor de Celsius para Fahrenheit**. |
|  `helpUrl`  |  string  |   Não  |  URL que fornece informações sobre a função. (Ela é exibida em um painel de tarefas.) Por exemplo, **http://contoso.com/help/convertcelsiustofahrenheit.html**. |
| `id`     | string | Sim | Um ID exclusivo da função. Este ID pode conter apenas caracteres alfanuméricos e pontos e não deve ser alterado depois de definido. |
|  `name`  |  string  |  Sim  |  O nome da função que os usuários finais veem no Excel. No Excel, esse nome de função será prefixado pelo namespace das funções personalizadas que é especificado no arquivo de manifesto XML. |
|  `options`  |  objeto  |  Não  |  Permite personalizar alguns aspectos de como e quando o Excel executa a função. Consulte o [objeto options](#options-object) para obter detalhes. |
|  `parameters`  |  matriz  |  Sim  |  Matriz que define os parâmetros de entrada para a função. Consulte a [matriz de parâmetros](#parameters-array) , para obter detalhes. |
|  `result`  |  objeto  |  Sim  |  Objeto que define o tipo de informação que é retornado pela função. Consulte o [objeto result](#result-object) para obter detalhes. |

## <a name="options"></a>options

O objeto  `options` permite personalizar alguns aspectos do como e quando o Excel executa a função. A tabela a seguir lista as propriedades do objeto `options`.

|  Propriedade  |  Tipo de dados  |  Obrigatório  |  Descrição  |
|:-----|:-----|:-----|:-----|
|  `cancelable`  |  booleano  |  Não<br/><br/>O valor padrão é `false`.  |  Se `true`, o Excel chama o manipulador de `onCanceled` sempre que o usuário realizar uma ação que tem o efeito de cancelar a função; por exemplo, disparando manualmente o recálculo ou editando uma célula referenciada pela função. Se você usar essa opção, o Excel chamará a função JavaScript com o parâmetro adicional `caller`. (***Não*** registre esse parâmetro na propriedade `parameters`). No corpo da função, um manipulador deve ser atribuído ao membro `caller.onCanceled`. Para saber mais, confira [Cancelar uma função](custom-functions-overview.md#canceling-a-function). |
|  `stream`  |  booleano  |  Não<br/><br/>O valor padrão é `false`.  |  Se `true`, a função pode atualizar o valor da célula repetidamente, mesmo quando invocada apenas uma vez. Essa opção é útil para alterar fontes de dados rapidamente, como um preço de ação. Se você usar essa opção, o Excel chamará a função JavaScript com o parâmetro adicional `caller`. (***Não*** registre esse parâmetro na propriedade `parameters` ). A função não deve ter a instrução `return`. Em vez disso, o valor de resultado é passado como o argumento do `caller.setResult` método de retorno de chamada. Para obter mais informações, consulte [Funções de fluxo contínuo](custom-functions-overview.md#streaming-functions). |

## <a name="parameters"></a>parameters

A propriedade  `parameters` é uma matriz de parâmetros. A tabela a seguir lista as propriedades de cada objeto.

|  Propriedade  |  Tipo de dados  |  Obrigatório  |  Descrição  |
|:-----|:-----|:-----|:-----|
|  `description`  |  string  |  Não |  Uma descrição do parâmetro.  |
|  `dimensionality`  |  string  |  Não  |  Deve ser **scalar** (um valor não-matriz) ou **matrix** (uma matriz bidimensional).  |
|  `name`  |  string  |  Sim  |  O nome do parâmetro. Esse nome é exibido no intelliSense do Excel.  |
|  `type`  |  string  |  Não  |  O tipo de dados do parâmetro. Deve ser **boolean**, **number** ou **string**.  |

## <a name="result"></a>result

O objeto  `results` define o tipo de informação retornado pela função. A tabela a seguir lista as propriedades do objeto `result`.

|  Propriedade  |  Tipo de dados  |  Obrigatório  |  Descrição  |
|:-----|:-----|:-----|:-----|
|  `dimensionality`  |  string  |  Não  |  Deve ser **scalar** (um valor não-matriz) ou **matrix** (uma matriz bidimensional). |
|  `type`  |  string  |  Sim  |  O tipo de dados do parâmetro. Deve ser **boolean**, **number** ou **string**.  |

## <a name="see-also"></a>Confira também

* [Criar funções personalizadas no Excel](custom-functions-overview.md)
* [Runtime para funções personalizadas do Excel](custom-functions-runtime.md)
* [Práticas recomendadas para funções personalizadas](custom-functions-best-practices.md)
* [Tutorial de funções personalizadas do Excel](excel-tutorial-custom-functions.md)
