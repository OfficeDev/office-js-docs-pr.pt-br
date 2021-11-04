---
title: Criar metadados JSON manualmente para funções personalizadas Excel
description: Defina os metadados JSON para funções personalizadas no Excel e associe sua ID de função e propriedades de nome.
ms.date: 11/03/2021
ms.localizationpriority: medium
ms.openlocfilehash: 28be374a88890d20294311599b06b16942edd9b7
ms.sourcegitcommit: ad5d7ab21f64012543fb2bd9226d90330d25468b
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 11/04/2021
ms.locfileid: "60749396"
---
# <a name="manually-create-json-metadata-for-custom-functions"></a>Criar metadados JSON manualmente para funções personalizadas

Conforme descrito no artigo visão geral de funções [personalizadas,](custom-functions-overview.md) um projeto de funções personalizadas deve incluir um arquivo de metadados JSON e um arquivo de script (JavaScript ou TypeScript) para registrar uma função, disponibilizando-a para uso. As funções personalizadas são registradas quando o usuário executa o complemento pela primeira vez e depois disso estão disponíveis para o mesmo usuário em todas as guias de trabalho.

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

Recomendamos usar a geração automática JSON quando possível, em vez de criar seu próprio arquivo JSON. A geração automática é menos propensa a erros do usuário e os arquivos `yo office` scaffolded já incluem isso. Para obter mais informações sobre marcas JSDoc e o processo de geração automática JSON, consulte [Autogenerate JSON metadata for custom functions](custom-functions-json-autogeneration.md).

No entanto, você pode fazer um projeto de funções personalizadas do zero. Esse processo exige que você:

- Escreva seu arquivo JSON.
- Verifique se o arquivo de manifesto está conectado ao arquivo JSON.
- Associe suas funções `id` e propriedades no arquivo de script para registrar suas `name` funções.

A imagem a seguir explica as diferenças entre usar `yo office` arquivos scaffold e escrever JSON do zero.

![Imagem das diferenças entre usar Yo Office e escrever seu próprio JSON.](../images/custom-functions-json.png)

> [!NOTE]
> Lembre-se de conectar seu manifesto ao arquivo JSON que você criar, por meio da seção em seu arquivo de manifesto XML se `<Resources>` você não usar o `yo office` gerador.

## <a name="authoring-metadata-and-connecting-to-the-manifest"></a>Autoria de metadados e conexão com o manifesto

Crie um arquivo JSON em seu projeto e forneça todos os detalhes sobre suas funções nele, como os parâmetros da função. Consulte o [exemplo de metadados a](#json-metadata-example) seguir e a referência de [metadados](#metadata-reference) para uma lista completa de propriedades de função.

Verifique se o arquivo de manifesto XML faz referência ao arquivo JSON na `<Resources>` seção, semelhante ao exemplo a seguir.

```json
<Resources>
    <bt:Urls>
        <bt:Url id="JSON-URL" DefaultValue="https://subdomain.contoso.com/config/customfunctions.json"/>
        <bt:Url id="JS-URL" DefaultValue="https://subdomain.contoso.com/dist/win32/ship/index.win32.bundle"/>
            <bt:Url id="HTML-URL" DefaultValue="https://subdomain.contoso.com/index.html"/>
    </bt:Urls>
    <bt:ShortStrings>
        <bt:String id="namespace" DefaultValue="CONTOSO"/>
    </bt:ShortStrings>
</Resources>
```

## <a name="json-metadata-example"></a>Exemplo de metadados JSON

O exemplo a seguir mostra o conteúdo de um arquivo de metadados JSON para um suplemento que define funções personalizadas. As seções que seguem este exemplo fornecem informações detalhadas sobre as propriedades individuais neste exemplo de JSON.

```json
{
  "allowCustomDataForDataTypeAny": true, // This property is currently only available in public preview.
  "allowErrorForDataTypeAny": true,
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
        "dimensionality": "scalar"
      },
      "parameters": []
    },
    {
      "id": "INCREMENTVALUE",
      "name": "INCREMENTVALUE",
      "description": "Count up from zero",
      "helpUrl": "http://www.contoso.com/help",
      "result": {
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
      "description": "Get the second highest number from a range",
      "helpUrl": "http://www.contoso.com/help",
      "result": {
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
> Um arquivo JSON de exemplo completo está disponível no [OfficeDev/Excel-Custom-Functions](https://github.com/OfficeDev/Excel-Custom-Functions/blob/77760adb1dcc53469183049bea08196734dbc114/config/customfunctions.json) GitHub histórico de confirmação do repositório. Como o projeto foi ajustado para gerar automaticamente o JSON, um exemplo completo de JSON manuscrito só está disponível em versões anteriores do projeto.

## <a name="metadata-reference"></a>Referência de metadados

### <a name="allowcustomdatafordatatypeany-preview"></a>allowCustomDataForDataTypeAny (visualização)

> [!NOTE]
> No momento, a propriedade está disponível na visualização pública e é compatível apenas com Office `allowCustomDataForDataTypeAny` no Windows. Os recursos de visualização estão sujeitas a alterações e não se destinam ao uso em um ambiente de produção. Recomendamos que você experimente apenas em ambiente de teste e desenvolvimento. Não use oa recursos de visualização em um ambiente de produção ou em documentos essenciais aos negócios.
>
> Para experimentar essa propriedade Office em Windows, você deve ter um número de com build Excel maior ou igual a 16.0.14623.20002. Para usar esse recurso, você precisa ingressar no [ Programa Office Insider](https://insider.office.com/) e, em seguida, escolher o nível Insider do **Canal beta**. Para saber mais, confira [ingressar no Programa Office Insider](https://insider.office.com/join/windows).

A `allowCustomDataForDataTypeAny` propriedade é um tipo de dados booleano. Definir esse valor para `true` permitir que uma função personalizada aceite tipos de dados como parâmetros e retorne valores. Para saber mais, confira [Conceitos básicos de funções e tipos de dados personalizados.](custom-functions-data-types-concepts.md)

Ao contrário da maioria das outras propriedades de metadados JSON, é uma propriedade de nível superior `allowCustomDataForDataTypeAny` e não contém sub-propriedades. Consulte o exemplo de código de [metadados JSON](#json-metadata-example) anterior para ver um exemplo de como formatar essa propriedade.

### <a name="allowerrorfordatatypeany"></a>allowErrorForDataTypeAny

A `allowErrorForDataTypeAny` propriedade é um tipo de dados booleano. Definir o valor para `true` permitir que uma função personalizada processe erros como valores de entrada. Todos os parâmetros com o tipo ou podem aceitar erros como valores `any` de entrada quando definido como `any[][]` `allowErrorForDataTypeAny` `true` . O valor `allowErrorForDataTypeAny` padrão é `false` .

> [!NOTE]
> Ao contrário das outras propriedades de metadados JSON, é uma propriedade de `allowErrorForDataTypeAny` nível superior e não contém sub-propriedades. Consulte o exemplo de código de [metadados JSON](#json-metadata-example) anterior para ver um exemplo de como formatar essa propriedade.

### <a name="functions"></a>functions

A propriedade `functions` é um conjunto de objetos de funções personalizadas. A tabela a seguir lista as propriedades de cada objeto.

| Propriedade      | Tipo de dados | Obrigatório | Descrição                                                                                                                                                                      |
| :------------ | :-------- | :------- | :------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `description` | string    | Não       | Descrição da função que é exibida aos usuários finais no Excel. Por exemplo, **Converte um valor em Celsius para Fahrenheit**.                                                            |
| `helpUrl`     | string    | Não       | A URL que fornece informações sobre a função. (Ela é exibida em um painel de tarefas). Por exemplo, `http://contoso.com/help/convertcelsiustofahrenheit.html`.                      |
| `id`          | string    | Sim      | Identificação exclusiva para a função. Essa ID pode conter apenas caracteres alfanuméricos e pontos e não deve ser alterada depois de configurada.                                            |
| `name`        | string    | Sim      | O nome da função que é exibida aos usuários finais no Excel. Em Excel, esse nome de função é prefixado pelo namespace de funções personalizadas especificado no arquivo de manifesto XML. |
| `options`     | object    | Não       | Permite que você personalize alguns aspectos de como e quando o Excel executa a função. Confira [opções](#options) para obter detalhes.                                                          |
| `parameters`  | array     | Sim      | Matriz que define os parâmetros de entrada para a função. Consulte [parâmetros](#parameters) para obter detalhes.                                                                             |
| `result`      | object    | Sim      | Objeto que define o tipo de informação que é retornada pela função do Excel. Confira [resultado](#result) para obter detalhes.                                                                 |

### <a name="options"></a>options

O objeto `options` permite que você personalize alguns aspectos de como e quando o Excel executa a função. A tabela a seguir lista as propriedades do objeto `options`.

| Propriedade          | Tipo de dados | Obrigatório                               | Descrição |
| :---------------- | :-------- | :------------------------------------- | :---------- |
| `cancelable`      | booliano   | Não<br/><br/>O valor padrão é `false`.  | Se o valor for `true`, o Excel chamará o manipulador `CancelableInvocation` sempre que o usuário realizar uma ação que tenha o efeito de cancelar a função, por exemplo, manualmente acionar um recálculo ou editar uma célula referenciada pela função. Funções canceláveis geralmente são usadas apenas para funções assíncronas que retornam um único resultado e precisam lidar com o cancelamento de uma solicitação de dados. Uma função não pode usar as `stream` propriedades `cancelable` e. |
| `requiresAddress` | booliano   | Não <br/><br/>O valor padrão é `false`. | Se `true` , sua função personalizada pode acessar o endereço da célula que a invocou. A `address` propriedade do parâmetro [invocação](custom-functions-parameter-options.md#invocation-parameter) contém o endereço da célula que invocou sua função personalizada. Uma função não pode usar as `stream` propriedades `requiresAddress` e. |
| `requiresParameterAddresses` | booliano   | Não <br/><br/>O valor padrão é `false`. | Se `true` , sua função personalizada pode acessar os endereços dos parâmetros de entrada da função. Essa propriedade deve ser usada em combinação com a propriedade do objeto de resultado e deve `dimensionality` ser definida como [](#result) `dimensionality` `matrix` . Consulte [Detectar o endereço de um parâmetro para](custom-functions-parameter-options.md#detect-the-address-of-a-parameter) obter mais informações. |
| `stream`          | booliano   | Não<br/><br/>O valor padrão é `false`.  | Se o valor for `true`, a função poderá gerar uma saída para a célula de forma repetida, mesmo quando invocada somente uma vez. Essa opção é útil para fontes de dados que mudam constantemente, como preços de ações. A função não deve ter instruções `return`. Em vez disso, o valor resultante é passado como o argumento do método de retorno `StreamingInvocation.setResult`. Para obter mais informações, consulte [Make a streaming function](custom-functions-web-reqs.md#make-a-streaming-function). |
| `volatile`        | booliano   | Não <br/><br/>O valor padrão é `false`. | Se , a função será recalculada sempre que Excel recalcular, em vez de somente quando os valores dependentes da fórmula `true` foram alterados. Uma função não pode usar as `stream` propriedades `volatile` e. Se as `stream` propriedades `volatile` e estão definidas como , a `true` propriedade volátil será ignorada. |

### <a name="parameters"></a>parâmetros

A propriedade `parameters` é uma matriz de objetos de parâmetro. A tabela a seguir lista as propriedades de cada objeto.

|  Propriedade  |  Tipo de dados  |  Obrigatório  |  Descrição  |
|:-----|:-----|:-----|:-----|
|  `description`  |  string  |  Não |  Uma descrição do parâmetro. Isso é exibido no Excel do IntelliSense.  |
|  `dimensionality`  |  string  |  Não  |  Deve ser `scalar` (um valor não matriz) `matrix` ou (uma matriz bidimensional).  |
|  `name`  |  string  |  Sim  |  O nome do parâmetro. Esse nome é exibido no Excel do IntelliSense.  |
|  `type`  |  string  |  Não  |  O tipo de dados do parâmetro. Pode ser , , ou , o que permite que você `boolean` use qualquer um dos três tipos `number` `string` `any` anteriores. Se essa propriedade não for especificada, o tipo de dados será padrão para `any` . |
|  `optional`  | booliano | Não | Se for `true`, o parâmetro será opcional. |
|`repeating`| booliano | Não | If `true` , parameters populate from a specified array. Observe que todas as funções de todos os parâmetros repetidos são consideradas parâmetros opcionais por definição.  |

### <a name="result"></a>result

O objeto `result` que define o tipo de informação que é retornado pela função. A tabela a seguir lista as propriedades do objeto `result`.

| Propriedade         | Tipo de dados | Obrigatório | Descrição                                                                          |
| :--------------- | :-------- | :------- | :----------------------------------------------------------------------------------- |
| `dimensionality` | string    | Não       | Deve ser `scalar` (um valor não matriz) `matrix` ou (uma matriz bidimensional). |
| `type` | string    | Não       | O tipo de dados do resultado. Pode ser , , ou (que permite que você `boolean` use qualquer um dos três tipos `number` `string` `any` anteriores). Se essa propriedade não for especificada, o tipo de dados será padrão para `any` . |

## <a name="associating-function-names-with-json-metadata"></a>Associar os nomes de função com metadados JSON

Para que uma função funcione corretamente, você precisa associar a propriedade da função `id` à implementação do JavaScript. Certifique-se de que haja uma associação, caso contrário, a função não será registrada e não será Excel. O exemplo de código a seguir mostra como fazer a associação usando o `CustomFunctions.associate()` método. A amostra define a função personalizada `add` e associa com o objeto no arquivo de metadados JSON onde o valor da `id` propriedade é **adicionar**.

```js
/**
 * Add two numbers
 * @customfunction
 * @param {number} first First number
 * @param {number} second Second number
 * @returns {number} The sum of the two numbers.
 */
function add(first, second) {
  return first + second;
}

CustomFunctions.associate("ADD", add);
```

O JSON a seguir mostra os metadados JSON associados à função personalizada anterior código JavaScript.

```json
{
  "functions": [
    {
      "description": "Add two numbers",
      "id": "ADD",
      "name": "ADD",
      "parameters": [
        {
          "description": "First number",
          "name": "first",
          "type": "number"
        },
        {
          "description": "Second number",
          "name": "second",
          "type": "number"
        }
      ],
      "result": {
        "type": "number"
      }
    }
  ]
}
```

Lembre-se das seguintes práticas recomendadas quando criar funções personalizadas no arquivo JavaScript e especificar as informações correspondentes no arquivo de metadados JSON.

- No arquivo de metadados JSON, verifique se o valor de cada propriedade `id` contém apenas caracteres alfanuméricos e pontos.

- No arquivo de metadados JSON, garanta que o valor de cada propriedade `id` seja exclusivo dentro do escopo do arquivo. Ou seja, nenhum objeto de duas funções no arquivo de metadados deve ter o mesmo valor `id`.

- Não altere o valor de uma propriedade `id` no arquivo de metadados JSON, depois de mapeá-lo para um nome de função JavaScript correspondente. Para alterar o nome da função que os usuários finais visualizam no Excel, atualize a propriedade `name` no arquivo de metadados JSON. No entanto, nunca altere o valor de uma propriedade `id` depois de estabelecida.

- No arquivo JavaScript, especifique uma associação de função personalizada usando `CustomFunctions.associate` após cada função.

O exemplo a seguir mostra os metadados JSON que correspondem às funções definidas no exemplo de código JavaScript anterior. Os `id` valores e propriedades estão em `name` maiúsculas, o que é uma prática prática prática ao descrever suas funções personalizadas. Você só precisará adicionar esse JSON se estiver preparando seu próprio arquivo JSON manualmente e não usando a geração automática. Para obter mais informações sobre a geração automática, consulte [Metadados JSON](custom-functions-json-autogeneration.md)de geração automática para funções personalizadas.

```json
{
  "$schema": "https://developer.microsoft.com/json-schemas/office-js/custom-functions.schema.json",
  "functions": [
    {
      "id": "ADD",
      "name": "ADD",
      ...
    },
    {
      "id": "INCREMENT",
      "name": "INCREMENT",
      ...
    }
  ]
}
```

## <a name="next-steps"></a>Próximas etapas

Aprenda as [práticas recomendadas para nomear sua](custom-functions-naming.md) função ou descubra como [localizar sua](custom-functions-localize.md) função usando o método JSON escrito à mão descrito anteriormente.

## <a name="see-also"></a>Confira também

- [Gerar metadados JSON automaticamente para funções personalizadas](custom-functions-json-autogeneration.md)
- [Opções de parâmetro de funções personalizadas](custom-functions-parameter-options.md)
- [Criar funções personalizadas no Excel](custom-functions-overview.md)
