---
ms.date: 09/20/2018
description: Defina metadados para funções personalizadas no Excel.
title: Metadados para funções personalizadas no Excel
ms.openlocfilehash: 815b0c6e65966867d9e5d953a40ffc705a63ee63
ms.sourcegitcommit: 470d8212b256275587e651abaa6f28beafebcab4
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 09/21/2018
ms.locfileid: "24062141"
---
# <a name="custom-functions-metadata"></a>Metadados de funções personalizadas

Quando você define [funções personalizadas](custom-functions-overview.md) dentro de seu suplemento do Excel, o seu projeto de suplemento deve incluir um arquivo de metadados JSON que fornece as informações que o Excel precisa para registrar as funções personalizadas e torná-las disponíveis para os usuários finais. Este artigo descreve o formato do arquivo JSON de metadados.

> [!NOTE]
> Para obter informações sobre os outros arquivos que você deve incluir em seu projeto de suplemento para habilitar funções personalizadas, confira [Criação de funções personalizadas no Excel](custom-functions-overview.md#learn-the-basics).

## <a name="example-metadata"></a>Exemplo de metadados

O exemplo a seguir mostra o conteúdo de um arquivo JSON de metadados para um suplemento que define funções personalizadas. As seções seguintes a esse exemplo fornecem informações detalhadas sobre as propriedades individuais deste exemplo JSON.

```json
{
    "functions": [
        {
            "id": "ADD42",
            "name": "ADD42",
            "description":  "Adds 42 to the input number",
            "helpUrl": "http://dev.office.com",
            "result": {
                "type": "number",
                "dimensionality": "scalar"
            },
            "parameters": [
                {
                    "name": "num",
                    "description": "Number",
                    "type": "number",
                    "dimensionality": "scalar"
                }
            ]
        },
        {
            "id": "ADD42ASYNC",
            "name": "ADD42ASYNC",
            "description":  "asynchronously wait 250ms, then add 42",
            "helpUrl": "http://dev.office.com",
            "result": {
                "type": "number",
                "dimensionality": "scalar"
            },
            "parameters": [
                {
                    "name": "num",
                    "description": "Number",
                    "type": "number",
                    "dimensionality": "scalar"
                }
            ]
        },
        {
            "id": "ISEVEN",
            "name": "ISEVEN", 
            "description":  "Determines whether a number is even",
            "helpUrl": "http://dev.office.com",
            "result": {
                "type": "boolean",
                "dimensionality": "scalar"
            },
            "parameters": [
                {
                    "name": "num",
                    "description": "the number to be evaluated",
                    "type": "number",
                    "dimensionality": "scalar"
                }
            ]
        },
        {
            "id": "GETDAY",
            "name": "GETDAY",
            "description": "Gets the day of the week",
            "helpUrl": "http://dev.office.com",
            "result": {
                "type": "string"
            },
            "parameters": []
        },
        {
            "id": "INCREMENTVALUE",
            "name": "INCREMENTVALUE", 
            "description":  "Counts up from zero",
            "helpUrl": "http://dev.office.com",
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
            "description":  "gets the second highest number from a range",
            "helpUrl": "http://dev.office.com",
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
|  `description`  |  sequência de caracteres  |  Não  |  Uma descrição da função que aparece na interface do usuário do Excel. Por exemplo, **Converte um valor Celsius em Fahrenheit**. |
|  `helpUrl`  |  sequência de caracteres  |   Não  |  A URL onde os usuários podem obter informações sobre a função. (É exibida em um painel de tarefas.) Por exemplo, **http://contoso.com/help/convertcelsiustofahrenheit.html**. |
| `id`     | sequência de caracteres | Sim | Um ID exclusivo para a função. Esse ID não deve ser alterado depois de ser definido. |
|  `name`  |  sequência de caracteres  |  Sim  |  O nome da função como será exibido (precedido de um namespace) na interface do usuário do Excel quando um usuário estiver selecionando uma função. Não precisa ser igual ao nome da função nos locais em que estiver definido no JavaScript. |
|  `options`  |  object  |  Não  |  Permite que você personalize alguns aspectos de como e quando o Excel executa a função. Confira [objeto options](#options-object) para obter detalhes. |
|  `parameters`  |  matriz  |  Sim  |  Matriz que define os parâmetros de entrada para a função. Confira [matriz de parâmetros](#parameters-array) para obter detalhes. |
|  `result`  |  object  |  Sim  |  Objeto que define o tipo de informação que é retornado pela função. Confira [objeto result](#result-object) para obter detalhes. |

## <a name="options"></a>options

O objeto `options` permite que você personalize alguns aspectos de como e quando o Excel executa a função. A tabela a seguir lista as propriedades do objeto `options`.

|  Propriedade  |  Tipo de dados  |  Obrigatório  |  Descrição  |
|:-----|:-----|:-----|:-----|
|  `cancelable`  |  booleano  |  Não, o padrão é `false`.  |  Se for `true`, o Excel chama o manipulador `onCanceled` sempre que o usuário executar uma ação que tenha o efeito de cancelar a função; por exemplo, ao disparar manualmente o recálculo ou ao editar uma célula referenciada pela função. Caso você use essa opção, o Excel chamará a função JavaScript com um parâmetro `caller` adicional. (***Não*** registre esse parâmetro na propriedade `parameters`). No corpo da função, um manipulador deve ser atribuído ao membro `caller.onCanceled`. Para obter mais informações, consulte [Cancelamento de uma função](custom-functions-overview.md#canceling-a-function). |
|  `stream`  |  booleano  |  Não, o padrão é `false`.  |  Se for `true`, a função pode ser repetidamente a saída da célula, mesmo quando invocada apenas uma vez. Essa opção é útil para fontes de dados que mudam rapidamente, como o preço de uma ação. Caso você use essa opção, o Excel chamará a função JavaScript com um parâmetro `caller` adicional. (***Não*** registre esse parâmetro na propriedade `parameters`). A função não deve ter a instrução `return`. Em vez disso, o valor do resultado é passado como argumento do método de retorno de chamada `caller.setResult`. Para obter mais informações, consulte [Funções de fluxo contínuo](custom-functions-overview.md#streamed-functions). |

## <a name="parameters"></a>parameters

A propriedade `parameters` é uma matriz de objetos de parâmetro. A tabela a seguir lista as propriedades de cada objeto.

|  Propriedade  |  Tipo de dados  |  Obrigatório  |  Descrição  |
|:-----|:-----|:-----|:-----|
|  `description`  |  sequência de caracteres  |  Não |  Uma descrição do parâmetro.  |
|  `dimensionality`  |  sequência de caracteres  |  Não  |  Deve ser **scalar** (um valor não-matriz) ou **matrix** (uma matriz bidimensional).  |
|  `name`  |  sequência de caracteres  |  Sim  |  O nome do parâmetro. Esse nome é exibido no IntelliSense do Excel.  |
|  `type`  |  sequência de caracteres  |  Não  |  O tipo de dados do parâmetro. Deve ser **boolean**, **number** ou **string**.  |

## <a name="result"></a>result

O objeto `results` define o tipo de informação que é retornado pela função. A tabela a seguir lista as propriedades do objeto `result`.

|  Propriedade  |  Tipo de dados  |  Obrigatório  |  Descrição  |
|:-----|:-----|:-----|:-----|
|  `dimensionality`  |  sequência de caracteres  |  Não  |  Deve ser **scalar** (um valor não-matriz) ou **matrix** (uma matriz bidimensional). |
|  `type`  |  sequência de caracteres  |  Sim  |  O tipo de dados do parâmetro. Deve ser **boolean**, **number** ou **string**.  |

## <a name="see-also"></a>Confira também

* [Criar funções personalizadas no Excel](custom-functions-overview.md)
* [Runtime para funções personalizadas do Excel](custom-functions-runtime.md)
* [Práticas recomendadas de funções personalizadas](custom-functions-best-practices.md)