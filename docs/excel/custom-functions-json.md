---
ms.date: 03/29/2019
description: Defina os metadados de funções personalizadas no Excel.
title: Metadados de funções personalizadas no Excel (visualização)
localization_priority: Normal
ms.openlocfilehash: 28a9a0207f7439af164eb9ca7c4b9ed9e966b3ed
ms.sourcegitcommit: 14ceac067e0e130869b861d289edb438b5e3eff9
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 04/04/2019
ms.locfileid: "31477548"
---
# <a name="custom-functions-metadata-preview"></a>Metadados de funções personalizadas (versão prévia)

Quando você define [funções personalizadas](custom-functions-overview.md) dentro de seu suplemento do Excel, o projeto do suplemento inclui um arquivo de metadados JSON que fornece as informações que o Excel requer para registrar as funções personalizadas e torná-las disponíveis para os usuários finais. Este arquivo é gerado:

- por você, em um arquivo JSON manuscrito
- nos comentários do JSDoc inseridos no início da função

As funções personalizadas são registradas quando o usuário executa o suplemento pela primeira vez e depois que eles estão disponíveis para o mesmo usuário em todas as pastas de trabalho.

Este artigo descreve o formato do arquivo de metadados JSON, supondo que você o esteja escrevendo à mão. Para obter informações sobre a geração de arquivo JSON de comentário JSDoc, consulte [GENERATE JSON Metadata for Custom Functions](custom-functions-json-autogeneration.md).

Para saber mais sobre outros arquivos que você deve incluir em seu projeto de suplemento para habilitar funções personalizadas, confira [Criar funções personalizadas no Excel](custom-functions-overview.md).

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

> Configurações do servidor no servidor que hospeda o arquivo JSON deve ter o [CORS](https://developer.mozilla.org/docs/Web/HTTP/CORS) habilitado para funções personalizadas funcionarem corretamente no Excel Online.

## <a name="example-metadata"></a>Exemplo de metadados

O exemplo a seguir mostra o conteúdo de um arquivo de metadados JSON para um suplemento que define funções personalizadas. As seções que seguem este exemplo fornecem informações detalhadas sobre as propriedades individuais neste exemplo de JSON.

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
> Um exemplo de arquivo JSON completo está disponível no repositório GitHub [OfficeDev/Excel-Custom-Functions](https://github.com/OfficeDev/Excel-Custom-Functions/blob/master/src/functions/functions.json).

## <a name="functions"></a>functions 

A propriedade `functions` é um conjunto de objetos de funções personalizadas. A tabela a seguir lista as propriedades de cada objeto.

|  Propriedade  |  Tipo de dados  |  Obrigatório  |  Descrição  |
|:-----|:-----|:-----|:-----|
|  `description`  |  string  |  Não  |  Descrição da função que é exibida aos usuários finais no Excel. Por exemplo, **Converte um valor em Celsius para Fahrenheit**. |
|  `helpUrl`  |  string  |   Não  |  A URL que fornece informações sobre a função. (Ela é exibida em um painel de tarefas). Por exemplo, **http://contoso.com/help/convertcelsiustofahrenheit.html**. |
| `id`     | string | Sim | Identificação exclusiva para a função. Essa ID pode conter apenas caracteres alfanuméricos e pontos e não deve ser alterada depois de configurada. |
|  `name`  |  string  |  Sim  |  O nome da função que é exibida aos usuários finais no Excel. No Excel, o nome da função será prefixado pelo namespace de funções personalizadas que é especificado no arquivo de manifesto XML. |
|  `options`  |  objeto  |  Não  |  Permite que você personalize alguns aspectos de como e quando o Excel executa a função. Confira [opções](#options) para obter detalhes. |
|  `parameters`  |  array  |  Sim  |  Matriz que define os parâmetros de entrada para a função. Confira os [parâmetros](#parameters) para obter detalhes. |
|  `result`  |  object  |  Sim  |  Objeto que define o tipo de informação que é retornada pela função do Excel. Confira [resultado](#result) para obter detalhes. |

## <a name="options"></a>options

O objeto `options` permite que você personalize alguns aspectos de como e quando o Excel executa a função. A tabela a seguir lista as propriedades do objeto `options`.

|  Propriedade  |  Tipo de dados  |  Obrigatório  |  Descrição  |
|:-----|:-----|:-----|:-----|
|  `cancelable`  |  booliano  |  Não<br/><br/>O valor padrão é `false`.  |  Se o valor for `true`, o Excel chamará o manipulador `onCanceled` sempre que o usuário realizar uma ação que tenha o efeito de cancelar a função, por exemplo, manualmente acionar um recálculo ou editar uma célula referenciada pela função. Se você usar essa opção, o Excel chamará a função JavaScript com um parâmetro `caller` adicional. (***Não*** registre este parâmetro na propriedade `parameters`). No corpo da função, um manipulador deve ser atribuído ao membro `caller.onCanceled`. Para saber mais, confira [Cancelar uma função](custom-functions-web-reqs.md#canceling-a-function). |
|  `stream`  |  booliano  |  Não<br/><br/>O valor padrão é `false`.  |  Se o valor for `true`, a função poderá gerar uma saída para a célula de forma repetida, mesmo quando invocada somente uma vez. Essa opção é útil para fontes de dados que mudam constantemente, como preços de ações. Se você usar essa opção, o Excel chamará a função JavaScript com um parâmetro `caller` adicional. (***Não*** registre este parâmetro na propriedade `parameters`). A função não deve ter instruções `return`. Em vez disso, o valor resultante é passado como o argumento do método de retorno `caller.setResult`. Para saber mais informações, confira [Funções de streaming](custom-functions-web-reqs.md#streaming-functions). |
|  `volatile`  | booliano | Não <br/><br/>O valor padrão é `false`. | <br /><br /> Se for `true`, a função será recalculada sempre que o Excel recalcular, em vez de apenas quando os valores dependentes da fórmula forem alterados. Uma função não pode ser de streaming e volátil ao mesmo tempo. Se as propriedades `stream` e `volatile` forem definidas como `true`, a opção volátil será ignorada. |

## <a name="parameters"></a>parâmetros

A propriedade `parameters` é uma matriz de objetos de parâmetro. A tabela a seguir lista as propriedades de cada objeto.

|  Propriedade  |  Tipo de dados  |  Obrigatório  |  Descrição  |
|:-----|:-----|:-----|:-----|
|  `description`  |  string  |  Não |  Uma descrição do parâmetro. Isso é exibido no IntelliSense do Excel.  |
|  `dimensionality`  |  string  |  Não  |  Deve ser **escalar** (um valor não matriz) ou **matriz** (uma matriz de 2 dimensões).  |
|  `name`  |  string  |  Sim  |  O nome do parâmetro. Esse nome é exibido no IntelliSense do Excel.  |
|  `type`  |  string  |  Não  |  O tipo de dados do parâmetro. Pode ser **booliano**, **número**, **cadeia de caracteres** ou **qualquer**, que permita usar qualquer um dos três tipos anteriores. Se essa propriedade não for especificada, o tipo de dados padrão será **qualquer**. |
|  `optional`  | booliano | Não | Se for `true`, o parâmetro será opcional. |

>[!NOTE]
> Se a propriedade `type` de um parâmetro opcional não for especificada ou definida como `any`, é provável que você tenha problemas, como erros de lint em seu IDE e parâmetros opcionais que não serão exibidos quando a função estiver sendo inserida em uma célula no Excel. A previsão é para ser alterado em dezembro de 2018.

## <a name="result"></a>result

O objeto `result` que define o tipo de informação que é retornado pela função. A tabela a seguir lista as propriedades do objeto `result`.

|  Propriedade  |  Tipo de dados  |  Obrigatório  |  Descrição  |
|:-----|:-----|:-----|:-----|
|  `dimensionality`  |  string  |  Não  |  Deve ser **escalar** (um valor não matriz) ou **matriz** (uma matriz de 2 dimensões). |
|  `type`  |  string  |  Sim  |  O tipo de dados do parâmetro. Deve ser **booliano**, **número**, **cadeia de caracteres** ou **qualquer**, que permita usar qualquer um dos três tipos anteriores. |

## <a name="see-also"></a>Confira também

* [Criar funções personalizadas no Excel](custom-functions-overview.md)
* [Tempo de execução de funções personalizadas do Excel](custom-functions-runtime.md)
* [Práticas recomendadas de funções personalizadas.](custom-functions-best-practices.md)
* [Log de alteração de funções personalizadas](custom-functions-changelog.md)
* [Tutorial de funções personalizadas do Excel](../tutorials/excel-tutorial-create-custom-functions.md)
