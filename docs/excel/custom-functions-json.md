---
title: Criar manualmente metadados JSON para funções personalizadas no Excel
description: Defina metadados JSON para funções personalizadas no Excel e associe suas propriedades de nome e ID de função.
ms.date: 12/28/2021
ms.localizationpriority: medium
ms.openlocfilehash: 2cd3b5266334e3397cd90fc24e29858250dfb284
ms.sourcegitcommit: b6a3815a1ad17f3522ca35247a3fd5d7105e174e
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 07/22/2022
ms.locfileid: "66958577"
---
# <a name="manually-create-json-metadata-for-custom-functions"></a>Criar manualmente metadados JSON para funções personalizadas

Conforme descrito no artigo de visão geral de funções personalizadas, um projeto de funções [personalizadas](custom-functions-overview.md) deve incluir um arquivo de metadados JSON e um arquivo de script (JavaScript ou TypeScript) para registrar uma função, disponibilizando-a para uso. As funções personalizadas são registradas quando o usuário executa o suplemento pela primeira vez e, depois disso, estão disponíveis para o mesmo usuário em todas as pastas de trabalho.

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

É recomendável usar a geração automática JSON quando possível, em vez de criar seu próprio arquivo JSON. A geração automática é menos propensa a erros do usuário e os `yo office` arquivos com scaffold já incluem isso. Para obter mais informações sobre marcas JSDoc e o processo de geração automática JSON, consulte Gerar automaticamente metadados [JSON para funções personalizadas](custom-functions-json-autogeneration.md).

No entanto, você pode criar um projeto de funções personalizadas do zero. Esse processo exige que você:

- Escreva seu arquivo JSON.
- Verifique se o arquivo de manifesto está conectado ao arquivo JSON.
- Associe suas funções `id` e propriedades `name` no arquivo de script para registrar suas funções.

A imagem a seguir explica as diferenças entre usar arquivos `yo office` scaffold e gravar JSON do zero.

![Imagem das diferenças entre usar o gerador Yeoman para Suplementos do Office e escrever seu próprio JSON.](../images/custom-functions-json.png)

> [!NOTE]
> Lembre-se de conectar o manifesto ao arquivo JSON que você criar, **\<Resources\>** por meio da seção no arquivo de manifesto XML se você não usar o gerador [Yeoman para suplementos do Office](../develop/yeoman-generator-overview.md).

## <a name="authoring-metadata-and-connecting-to-the-manifest"></a>Criando metadados e conectando-se ao manifesto

Crie um arquivo JSON em seu projeto e forneça todos os detalhes sobre suas funções nele, como os parâmetros da função. Consulte o [exemplo de metadados a seguir](#json-metadata-example) e [a referência de metadados](#metadata-reference) para obter uma lista completa de propriedades de função.

Verifique se o arquivo de manifesto XML faz referência ao arquivo JSON **\<Resources\>** na seção, semelhante ao exemplo a seguir.

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
> Um arquivo JSON de exemplo completo está disponível no histórico de confirmação do repositório [GitHub OfficeDev/Excel-Custom-Functions](https://github.com/OfficeDev/Excel-Custom-Functions/blob/77760adb1dcc53469183049bea08196734dbc114/config/customfunctions.json) . Como o projeto foi ajustado para gerar automaticamente JSON, uma amostra completa de JSON manuscrito só está disponível em versões anteriores do projeto.

## <a name="metadata-reference"></a>Referência de metadados

### <a name="allowcustomdatafordatatypeany-preview"></a>allowCustomDataForDataTypeAny (versão prévia)

> [!NOTE]
> Atualmente `allowCustomDataForDataTypeAny` , a propriedade está disponível em versão prévia pública e é compatível apenas com o Office no Windows. Os recursos de visualização estão sujeitas a alterações e não se destinam ao uso em um ambiente de produção. Recomendamos que você experimente apenas em ambiente de teste e desenvolvimento. Não use oa recursos de visualização em um ambiente de produção ou em documentos essenciais aos negócios.
>
> Para experimentar essa propriedade no Office no Windows, você deve ter um número de build do Excel maior ou igual a 16.0.14623.20002. Para usar esse recurso, você precisa ingressar no [ Programa Office Insider](https://insider.office.com/) e, em seguida, escolher o nível Insider do **Canal beta**. Para saber mais, confira [ingressar no Programa Office Insider](https://insider.office.com/join/windows).

A `allowCustomDataForDataTypeAny` propriedade é um tipo de dados booliano. Definir esse valor para permitir `true` que uma função personalizada aceite tipos de dados como parâmetros e valores retornados. Para saber mais, confira [Funções personalizadas e tipos de dados](custom-functions-data-types-concepts.md).

Ao contrário da maioria das outras propriedades de metadados JSON, `allowCustomDataForDataTypeAny` é uma propriedade de nível superior e não contém sub-propriedades. Consulte o exemplo de código [de metadados JSON](#json-metadata-example) anterior para obter um exemplo de como formatar essa propriedade.

### <a name="allowerrorfordatatypeany"></a>allowErrorForDataTypeAny

A `allowErrorForDataTypeAny` propriedade é um tipo de dados booliano. Definir o valor para permitir `true` que uma função personalizada processe erros como valores de entrada. Todos os parâmetros com o tipo ou `any` podem `any[][]` aceitar erros como valores de entrada quando `allowErrorForDataTypeAny` são definidos como `true`. O valor `allowErrorForDataTypeAny` padrão é `false`.

> [!NOTE]
> Ao contrário das outras propriedades de metadados JSON, `allowErrorForDataTypeAny` é uma propriedade de nível superior e não contém sub-propriedades. Consulte o exemplo de código [de metadados JSON](#json-metadata-example) anterior para obter um exemplo de como formatar essa propriedade.

### <a name="functions"></a>functions

A propriedade `functions` é um conjunto de objetos de funções personalizadas. A tabela a seguir lista as propriedades de cada objeto.

| Propriedade      | Tipo de dados | Obrigatório | Descrição                                                                                                                                                                      |
| :------------ | :-------- | :------- | :------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `description` | string    | Não       | Descrição da função que é exibida aos usuários finais no Excel. Por exemplo, **Converte um valor em Celsius para Fahrenheit**.                                                            |
| `helpUrl`     | string    | Não       | A URL que fornece informações sobre a função. (Ela é exibida em um painel de tarefas). Por exemplo, `http://contoso.com/help/convertcelsiustofahrenheit.html`.                      |
| `id`          | string    | Sim      | Identificação exclusiva para a função. Essa ID pode conter apenas caracteres alfanuméricos e pontos e não deve ser alterada depois de configurada.                                            |
| `name`        | string    | Sim      | O nome da função que é exibida aos usuários finais no Excel. No Excel, esse nome de função é prefixado pelo namespace de funções personalizadas especificado no arquivo de manifesto XML. |
| `options`     | object    | Não       | Permite que você personalize alguns aspectos de como e quando o Excel executa a função. Confira [opções](#options) para obter detalhes.                                                          |
| `parameters`  | array     | Sim      | Matriz que define os parâmetros de entrada para a função. Consulte [os parâmetros](#parameters) para obter detalhes.                                                                             |
| `result`      | object    | Sim      | Objeto que define o tipo de informação que é retornada pela função do Excel. Confira [resultado](#result) para obter detalhes.                                                                 |

### <a name="options"></a>options

O objeto `options` permite que você personalize alguns aspectos de como e quando o Excel executa a função. A tabela a seguir lista as propriedades do objeto `options`.

| Propriedade          | Tipo de dados | Obrigatório                               | Descrição |
| :---------------- | :-------- | :------------------------------------- | :---------- |
| `cancelable`      | booliano   | Não<br/><br/>O valor padrão é `false`.  | Se o valor for `true`, o Excel chamará o manipulador `CancelableInvocation` sempre que o usuário realizar uma ação que tenha o efeito de cancelar a função, por exemplo, manualmente acionar um recálculo ou editar uma célula referenciada pela função. As funções canceláveis normalmente são usadas apenas para funções assíncronas que retornam um único resultado e precisam lidar com o cancelamento de uma solicitação de dados. Uma função não pode usar as propriedades `stream` e as `cancelable` propriedades. |
| `requiresAddress` | booliano   | Não <br/><br/>O valor padrão é `false`. | If `true`, sua função personalizada pode acessar o endereço da célula que a invocou. A `address` propriedade do parâmetro [de invocação](custom-functions-parameter-options.md#invocation-parameter) contém o endereço da célula que invocou sua função personalizada. Uma função não pode usar as propriedades `stream` e as `requiresAddress` propriedades. |
| `requiresParameterAddresses` | booliano   | Não <br/><br/>O valor padrão é `false`. | If `true`, sua função personalizada pode acessar os endereços dos parâmetros de entrada da função. Essa propriedade deve ser usada em combinação com a `dimensionality` propriedade do objeto [de](#result) resultado e `dimensionality` deve ser definida como `matrix`. Consulte [Detectar o endereço de um parâmetro para](custom-functions-parameter-options.md#detect-the-address-of-a-parameter) obter mais informações. |
| `stream`          | booliano   | Não<br/><br/>O valor padrão é `false`.  | Se o valor for `true`, a função poderá gerar uma saída para a célula de forma repetida, mesmo quando invocada somente uma vez. Essa opção é útil para fontes de dados que mudam constantemente, como preços de ações. A função não deve ter instruções `return`. Em vez disso, o valor do resultado é passado como o argumento da função `StreamingInvocation.setResult` de retorno de chamada. Para obter mais informações, consulte [Criar uma função de streaming](custom-functions-web-reqs.md#make-a-streaming-function). |
| `volatile`        | booliano   | Não <br/><br/>O valor padrão é `false`. | Se `true`, a função será recalculada sempre que o Excel for recalculado, em vez de somente quando os valores dependentes da fórmula tiverem sido alterados. Uma função não pode usar as propriedades `stream` e as `volatile` propriedades. Se a `stream` propriedade e `volatile` as propriedades forem definidas como `true`, a propriedade volátil será ignorada. |

### <a name="parameters"></a>parâmetros

A propriedade `parameters` é uma matriz de objetos de parâmetro. A tabela a seguir lista as propriedades de cada objeto.

|  Propriedade  |  Tipo de dados  |  Obrigatório  |  Descrição  |
|:-----|:-----|:-----|:-----|
|  `description`  |  string  |  Não |  Uma descrição do parâmetro. Isso é exibido no IntelliSense do Excel.  |
|  `dimensionality`  |  string  |  Não  |  Deve ser `scalar` (um valor não matriz) ou `matrix` (uma matriz bidimensional).  |
|  `name`  |  string  |  Sim  |  O nome do parâmetro. Esse nome é exibido no IntelliSense do Excel.  |
|  `type`  |  string  |  Não  |  O tipo de dados do parâmetro. Pode ser `boolean`, `number``string`ou `any`, o que permite que você use qualquer um dos três tipos anteriores. Se essa propriedade não for especificada, o tipo de dados padrão será `any`. |
|  `optional`  | booliano | Não | Se for `true`, o parâmetro será opcional. |
|`repeating`| booliano | Não | Se `true`, os parâmetros são preenchidos de uma matriz especificada. Observe que todas as funções que repetem parâmetros são consideradas parâmetros opcionais por definição.  |

### <a name="result"></a>result

O objeto `result` que define o tipo de informação que é retornado pela função. A tabela a seguir lista as propriedades do objeto `result`.

| Propriedade         | Tipo de dados | Obrigatório | Descrição                                                                          |
| :--------------- | :-------- | :------- | :----------------------------------------------------------------------------------- |
| `dimensionality` | string    | Não       | Deve ser `scalar` (um valor não matriz) ou `matrix` (uma matriz bidimensional). |
| `type` | string    | Não       | O tipo de dados do resultado. Pode ser `boolean`, `number``string`, ou `any` (que permite que você use qualquer um dos três tipos anteriores). Se essa propriedade não for especificada, o tipo de dados padrão será `any`. |

## <a name="associating-function-names-with-json-metadata"></a>Associar os nomes de função com metadados JSON

Para que uma função funcione corretamente, você precisa associar a propriedade `id` da função à implementação do JavaScript. Verifique se há uma associação; caso contrário, a função não será registrada e não poderá ser usada no Excel. O exemplo de código a seguir mostra como fazer a associação usando a `CustomFunctions.associate()` função. A amostra define a função personalizada `add` e associa com o objeto no arquivo de metadados JSON onde o valor da `id` propriedade é **adicionar**.

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

O JSON a seguir mostra os metadados JSON associados ao código JavaScript da função personalizada anterior.

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

- No arquivo JavaScript, especifique uma associação de função personalizada usando após `CustomFunctions.associate` cada função.

O exemplo a seguir mostra os metadados JSON que correspondem às funções definidas no exemplo de código JavaScript anterior. Os `id` valores `name` e a propriedade estão em letras maiúsculas, o que é uma prática recomendada ao descrever suas funções personalizadas. Você só precisará adicionar esse JSON se estiver preparando seu próprio arquivo JSON manualmente e não estiver usando a geração automática. Para obter mais informações sobre a geração automática, consulte [Geração automática de metadados JSON para funções personalizadas](custom-functions-json-autogeneration.md).

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

Conheça as [práticas recomendadas para nomear sua](custom-functions-naming.md) função ou descubra como [localizar](custom-functions-localize.md) sua função usando o método JSON manuscrito descrito anteriormente.

## <a name="see-also"></a>Confira também

- [Gerar metadados JSON automaticamente para funções personalizadas](custom-functions-json-autogeneration.md)
- [Opções de parâmetro de funções personalizadas](custom-functions-parameter-options.md)
- [Criar funções personalizadas no Excel](custom-functions-overview.md)
