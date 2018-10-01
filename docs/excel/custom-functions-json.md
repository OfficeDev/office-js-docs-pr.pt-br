---
ms.date: 09/27/2018
description: Defina metadados para funções personalizadas no Excel.
title: Metadados para funções personalizadas no Excel
ms.openlocfilehash: a179a9c4bc071200cab1377c5e48913bfc8358cf
ms.sourcegitcommit: 1852ae367de53deb91d03ca55d16eb69709340d3
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 09/29/2018
ms.locfileid: "25348791"
---
# <a name="custom-functions-metadata-preview"></a>Metadados de funções personalizadas (versão prévia)

Quando você define [funções personalizadas](custom-functions-overview.md) no suplemento do Excel, o projeto de suplemento deve incluir um arquivo de metadados JSON que forneça as informações necessárias para o Excel registrar as funções personalizadas e disponibilizá-las aos usuários finais. Este artigo descreve o formato do arquivo de metadados JSON.

Para obter informações sobre os outros arquivos que você deve incluir no projeto de suplemento para habilitar funções personalizadas, confira [Criar funções personalizadas no Excel](custom-functions-overview.md).

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

## <a name="example-metadata"></a>Exemplo de metadados

O exemplo a seguir mostra o conteúdo de um arquivo JSON de metadados para um suplemento que define funções personalizadas. As seções que seguem este exemplo fornecem informações detalhadas sobre as propriedades individuais nesse exemplo de JSON.

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
> Um exemplo completo de arquivo JSON está disponível no [repositório GitHub OfficeDev/Excel-Custom-Functions](https://github.com/OfficeDev/Excel-Custom-Functions/blob/master/config/customfunctions.json).

## <a name="functions"></a>functions 

A propriedade `functions` é uma matriz de objetos de funções personalizadas. A tabela a seguir lista as propriedades de cada objeto.

|  Propriedade  |  Tipo de dados  |  Obrigatório  |  Descrição  |
|:-----|:-----|:-----|:-----|
|  `description`  |  sequência de caracteres  |  Não  |  A descrição da função que os usuários finais veem no Excel. Por exemplo, **Converte um valor de Celsius para Fahrenheit**. |
|  `helpUrl`  |  sequência de caracteres  |   Não  |  URL que fornece informações sobre a função. (Ela é exibida em um painel de tarefas.) Por exemplo, **http://contoso.com/help/convertcelsiustofahrenheit.html**. |
| `id`     | sequência de caracteres | Sim | Um ID exclusivo para a função. Esse ID não deve ser alterado depois de ser definido. |
|  `name`  |  sequência de caracteres  |  Sim  |  O nome da função que os usuários finais veem no Excel. No Excel, esse nome de função terá como prefixo o namespace das funções personalizadas especificado no arquivo de manifesto XML. |
|  `options`  |  object  |  Não  |  Permite que você personalize alguns aspectos de como e quando o Excel executa a função. Confira [objeto options](#options-object) para obter detalhes. |
|  `parameters`  |  matriz  |  Sim  |  Matriz que define os parâmetros de entrada para a função. Confira [matriz de parâmetros](#parameters-array) para obter detalhes. |
|  `result`  |  objeto  |  Sim  |  Objeto que define o tipo de informação que é retornado pela função. Confira [objeto result](#result-object) para obter detalhes. |

## <a name="options"></a>options

O objeto `options` permite que você personalize alguns aspectos de como e quando o Excel executa a função. A tabela a seguir lista as propriedades do objeto `options`.

|  Propriedade  |  Tipo de dados  |  Obrigatório  |  Descrição  |
|:-----|:-----|:-----|:-----|
|  `cancelable`  |  booleano  |  Não<br/><br/>O valor padrão é `false`.  |  Se for `true`, o Excel chama o manipulador `onCanceled` sempre que o usuário executar uma ação que tenha o efeito de cancelar a função; por exemplo, acionando manualmente o recálculo ou editando uma célula referenciada pela função. Caso você use essa opção, o Excel chamará a função JavaScript com um parâmetro `caller` adicional. (***Não*** registre esse parâmetro na propriedade `parameters`). No corpo da função, um manipulador deve ser atribuído ao membro `caller.onCanceled`. Para obter mais informações, consulte [Cancelamento de uma função](custom-functions-overview.md#canceling-a-function). |
|  `stream`  |  booleano  |  Não<br/><br/>O valor padrão é `false`.  |  Se for `true`, a função pode modificar o valor da célula repetidamente, mesmo quando invocada apenas uma vez. Essa opção é útil para fontes de dados que mudam rapidamente, como o preço de uma ação. Caso você use essa opção, o Excel chamará a função JavaScript com um parâmetro `caller` adicional. (***Não*** registre esse parâmetro na propriedade `parameters`). A função não deve ter a instrução `return`. Em vez disso, o valor do resultado é passado como argumento do método de retorno de chamada `caller.setResult`. Para obter mais informações, consulte [Funções de fluxo contínuo](custom-functions-overview.md#streamed-functions). |

## <a name="parameters"></a>parameters

A propriedade `parameters` é uma matriz de objetos de parâmetro. A tabela a seguir lista as propriedades de cada objeto.

|  Propriedade  |  Tipo de dados  |  Obrigatório  |  Descrição  |
|:-----|:-----|:-----|:-----|
|  `description`  |  sequência de caracteres  |  Não |  Uma descrição do parâmetro.  |
|  `dimensionality`  |  sequência de caracteres  |  Não  |  Deve ser **scalar** (um valor não-matriz) ou **matrix** (uma matriz bidimensional).  |
|  `name`  |  sequência de caracteres  |  Sim  |  O nome do parâmetro. Esse nome é exibido no IntelliSense do Excel.  |
|  `type`  |  sequência de caracteres  |  Não  |  O tipo de dado do parâmetro. Deve ser **boolean**, **number**ou **string**.  |

## <a name="result"></a>result

O objeto `results` define o tipo de informação que é retornado pela função. A tabela a seguir lista as propriedades do objeto `result`.

|  Propriedade  |  Tipo de dados  |  Obrigatório  |  Descrição  |
|:-----|:-----|:-----|:-----|
|  `dimensionality`  |  sequência de caracteres  |  Não  |  Deve ser **scalar** (um valor não-matriz) ou **matrix** (uma matriz bidimensional). |
|  `type`  |  sequência de caracteres  |  Sim  |  O tipo de dados do parâmetro. Deve ser **boolean**, **number** ou **string**.  |

## <a name="see-also"></a>Confira também

* [Criar funções personalizadas no Excel](custom-functions-overview.md)
* [Runtime para funções personalizadas do Excel](custom-functions-runtime.md)
* [Práticas recomendadas para funções personalizadas](custom-functions-best-practices.md)
* [Tutorial de funções personalizadas do Excel](excel-tutorial-custom-functions.md)