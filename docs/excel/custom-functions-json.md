# <a name="custom-function-metadata"></a>Metadados da função personalizada

Ao incluir [funções personalizadas](custom-functions-overview.md) em um suplemento do Excel, você deve hospedar um arquivo JSON que contenha metadados sobre as funções (além de hospedar um arquivo JavaScript com as funções e um arquivo HTML sem interface do usuário para servir como pai do arquivo JavaScript). Este artigo descreve o formato do arquivo JSON com exemplos.

Há um arquivo JSON de amostra completo disponível [aqui](https://github.com/OfficeDev/Excel-Custom-Functions/blob/master/customfunctions.json).

## <a name="functions-array"></a>Matriz de funções

Os metadados são um objeto JSON que contém uma única propriedade `functions` cujo valor é uma matriz de objetos. Cada um desses objetos representa uma função personalizada. A tabela a seguir contém suas propriedades:

|  Propriedade  |  Tipo de dados  |  Obrigatório?  |  Descrição  |
|:-----|:-----|:-----|:-----|
|  `description`  |  sequência de caracteres  |  Não  |  Uma descrição da função que aparece na interface do usuário do Excel. Por exemplo, “Converte um valor Celsius em Fahrenheit”. |
|  `helpUrl`  |  sequência de caracteres  |   Não  |  A URL na qual seus usuários podem obter ajuda sobre a função. (Ela é exibida em um painel de tarefas.) Por exemplo, “http://contoso.com/help/convertcelsiustofahrenheit.html”  |
|  `name`  |  sequência de caracteres  |  Sim  |  O nome da função como será exibido (precedido de um namespace) na interface do usuário do Excel quando um usuário estiver selecionando uma função. Deve ser o mesmo que o nome da função nos locais em que estiver definido no JavaScript. |
|  `options`  |  objeto  |  Não  |  Configurar como o Excel processa a função. Consulte [objeto de opções](#options-object) para obter detalhes. |
|  `parameters`  |  matriz  |  Sim  |  Metadados sobre os parâmetros para a função. Consulte [matriz de parâmetros](#parameters-array) para obter detalhes. |
|  `result`  |  objeto  |  Sim  |  Metadados sobre o valor retornado pela função. Consulte [objeto de resultado](#result-object) para obter detalhes. |

## <a name="options-object"></a>Objeto de opções

O objeto `options` configura como o Excel processa a função. A tabela a seguir contém suas propriedades:

|  Propriedade  |  Tipo de dados  |  Obrigatório?  |  Descrição  |
|:-----|:-----|:-----|:-----|
|  `cancelable`  |  booleano  |  Não, o padrão é `false`.  |  Se for `true`, o Excel chama o manipulador `onCanceled` sempre que o usuário executar uma ação que tenha o efeito de cancelar a função; por exemplo, ao disparar manualmente o recálculo ou ao editar uma célula referenciada pela função. Caso você use essa opção, o Excel chamará a função JavaScript com um parâmetro `caller` adicional. (***Não*** registre esse parâmetro na propriedade `parameters`). No corpo da função, um manipulador deve ser atribuído ao membro `caller.onCanceled`. Observe que `cancelable` e `sync` não podem ambos ser `true`.  |
|  `stream`  |  booleano  |  Não, o padrão é `false`.  |  Se for `true`, a função pode ser repetidamente a saída da célula, mesmo quando invocada apenas uma vez. Essa opção é útil para fontes de dados que mudam rapidamente, como o preço de uma ação. Caso você use essa opção, o Excel chamará a função JavaScript com um parâmetro `caller` adicional. (***Não*** registre esse parâmetro na propriedade `parameters`). A função não deve ter a instrução `return`. Em vez disso, o valor do resultado é passado como o argumento do método de retorno de chamada `caller.setResult`. Observe que `stream` e `sync` não podem ambos ser `true`.|
|  `sync`  |  booleano  |  Não, o padrão é `false`  |  Se for `true`, a função é executada de forma síncrona e deve retornar um valor. Se for `false`, a função é executada de forma assíncrona e deve retornar o objeto `OfficeExtension.Promise`. Observe que talvez `sync` não seja `true` se `cancelable` ou `stream` for `true`.  |

## <a name="parameters-array"></a>Matriz de parâmetros

A propriedade `parameters` é uma matriz de objetos. Cada um desses objetos representa um parâmetro. A tabela a seguir contém suas propriedades:

|  Propriedade  |  Tipo de dados  |  Obrigatório?  |  Descrição  |
|:-----|:-----|:-----|:-----|
|  `description`  |  sequência de caracteres  |  Não |  Uma descrição do parâmetro.  |
|  `dimensionality`  |  sequência de caracteres  |  Sim  |  Deve ser “escalar”, ou seja, um valor que não é matriz; ou “matriz”, ou seja, uma matriz de matrizes de linhas.  |
|  `name`  |  sequência de caracteres  |  Sim  |  O nome do parâmetro. Esse nome é exibido no IntelliSense do Excel.  |
|  `type`  |  sequência de caracteres  |  Sim  |  O tipo de dados do parâmetro. Deve ser “booleano”, “número” ou “sequência de caracteres”.  |

## <a name="result-object"></a>Objeto de resultado

A propriedade `results` fornece metadados sobre o valor retornado da função. A tabela a seguir contém suas propriedades:

|  Propriedade  |  Tipo de dados  |  Obrigatório?  |  Descrição  |
|:-----|:-----|:-----|:-----|
|  `dimensionality`  |  sequência de caracteres  |  Não  |  Deve ser “escalar”, ou seja, um valor que não é matriz; ou “matriz”, ou seja, uma matriz de matrizes de linhas.  |
|  `type`  |  sequência de caracteres  |  Sim  |  O tipo de dados do parâmetro. Deve ser “booleano”, “número” ou “sequência de caracteres”.  |

## <a name="example"></a>Exemplo

O código JSON a seguir é um exemplo de um arquivo de metadados para funções personalizadas.

```json
{
    "functions": [
        {
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
            ],
            "options": {
                "sync": true
            }
        },
        {
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
            ],
            "options": {
                "sync": false
            }
        },
        {
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
            ],
            "options": {
                "sync": true
            }
        },
        {
            "name": "GETDAY",
            "description": "Gets the day of the week",
            "helpUrl": "http://dev.office.com",
            "result": {
                "type": "string"
            },
            "parameters": [],
            "options": {
                "sync": true
            }
        },
        {
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
                "sync": false,
                "stream": true,
                "cancelable": true
            }
        },
        {
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
            ],
            "options": {
                "sync": true
            }
        }
    ]
}

```

## <a name="see-also"></a>Confira também
[Funções personalizadas](custom-functions-overview.md)<br>
[Diretrizes e exemplos de fórmulas de matriz](https://support.office.com/article/Guidelines-and-examples-of-array-formulas-7d94a64e-3ff3-4686-9372-ecfd5caa57c7)
