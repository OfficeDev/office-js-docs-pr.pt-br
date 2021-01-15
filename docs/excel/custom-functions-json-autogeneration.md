---
ms.date: 12/22/2020
description: Use tags JSDoc para criar dinamicamente seus metadados JSON de funções personalizadas.
title: Gerar metadados JSON automaticamente para funções personalizadas
localization_priority: Normal
ms.openlocfilehash: 4f7553646f51518fd4c5ed6c66081f910ec813a1
ms.sourcegitcommit: 48b9c3b63668b2a53ce73f92ce124ca07c5ca68c
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 12/28/2020
ms.locfileid: "49735554"
---
# <a name="autogenerate-json-metadata-for-custom-functions"></a>Gerar metadados JSON automaticamente para funções personalizadas

Quando uma função personalizada do Excel é gravada em JavaScript ou em TypeScript, as [marcações JSDoc](https://jsdoc.app/) são usadas para fornecer informações adicionais sobre a função personalizada. As marcações JSDoc são usadas no momento da criação para criar o arquivo de metadados JSON. O uso de marcas JSDoc poupa você do esforço de [editar manualmente o arquivo de metadados JSON](custom-functions-json.md).

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

Adicione a marcação `@customfunction` nos comentários de código de uma função JavaScript ou TypeScript para marcá-la como uma função personalizada.

Os tipos de parâmetros da função podem ser fornecidos usando a marcação [@param](#param) em JavaScript ou do [Tipo de função](https://www.typescriptlang.org/docs/handbook/functions.html) em TypeScript. Para saber mais, confira a marcação [@param](#param) e as seções [Tipos](#types).

### <a name="adding-a-description-to-a-function"></a>Adicionando uma descrição a uma função

A descrição é exibida para o usuário como texto de ajuda quando eles precisam de ajuda para entender o que a função personalizada executa. A descrição não requer nenhuma tag específica. Basta digitar uma breve descrição de texto no comentário JSDoc. Em geral, a descrição é colocada no início da seção de comentários do JSDoc, mas funcionará independentemente de onde seja colocada.

Para ver exemplos das descrições de funções internas, abra o Excel, vá para a guia **Fórmulas** e escolha **Inserir função**. Você pode navegar por todas as descrições de funções e também ver suas próprias funções personalizadas listadas.

No exemplo a seguir, a frase "Calcula o volume de uma esfera." é a descrição da função personalizada.

```js
/**
/* Calculates the volume of a sphere.
/* @customfunction VOLUME
...
 */
```


## <a name="jsdoc-tags"></a>Marcações JSDoc

As seguintes marcas JSDoc são suportadas nas funções personalizadas do Excel.

* [@cancelable](#cancelable)
* [@customfunction](#customfunction) nome de identificação
* [@helpurl](#helpurl) url
* [@param](#param) _{type}_ nome e descrição
* [@requiresAddress](#requiresAddress)
* [@returns](#returns) _{type}_
* [@streaming](#streaming)
* [@volatile](#volatile)

---
<a id="cancelable"></a>

### <a name="cancelable"></a>@cancelable

Indica que uma função personalizada executa uma ação quando a função é cancelada.

O último parâmetro da função deve ser do tipo `CustomFunctions.CancelableInvocation`. A função pode atribuir uma função à `oncanceled` propriedade para indicar o resultado quando a função é cancelada.

Se o último parâmetro da função for do tipo `CustomFunctions.CancelableInvocation`, ela será considerada `@cancelable`, mesmo se a tag não estiver presente.

Uma função não pode ter as tags `@cancelable` e `@streaming` ao mesmo tempo.

<a id="customfunction"></a>

### <a name="customfunction"></a>@customfunction

Sintaxe: @customfunction _id_ _nome_

Essa marca indica que a função JavaScript/TypeScript é uma função personalizada do Excel. É necessário criar metadados para a função personalizada.

Veja a seguir um exemplo dessa marca.

```js
/**
 * Increments a value once a second.
 * @customfunction
 * ...
 */
```

#### <a name="id"></a>id

O `id` identifica uma função personalizada.

* Se `id` não for fornecido, o nome da função JavaScript/TypeScript será convertido em maiúsculas e os caracteres não permitidos serão removidos.
* O `id` deve ser exclusivo para todas as funções personalizadas.
* Os caracteres permitidos estão limitados a: A-Z, a-z, 0-9, sublinhados (\_) e ponto (.).

No exemplo a seguir, incremento é o `id` e o `name` da função.

```js
/**
 * Increments a value once a second.
 * @customfunction INCREMENT
 * ...
 */
```

#### <a name="name"></a>nome

Fornece a exibição `name` da função personalizada.

* Se o nome não for fornecido, o id também será usado como nome.
* Caracteres permitidos: Letras de [caractere Alfabético Unicode](https://www.unicode.org/reports/tr44/tr44-22.html#Alphabetic), números, ponto (.) e sublinhado (\_).
* Deve começar com uma letra.
* O comprimento máximo é de 128 caracteres.

No exemplo a seguir, Inc é a `id`da função e `increment` é o `name`.

```js
/**
 * Increments a value once a second.
 * @customfunction INC INCREMENT
 * ...
 */
```

### <a name="description"></a>descrição

Uma descrição aparece para os usuários no Excel à medida que estão inserindo a função e especifica o que a função faz. Uma descrição não exige nenhuma tag específica. Adicione uma descrição a uma função personalizada acrescentando uma frase para descrever o que a função realiza dentro do comentário JSDoc. Por padrão, qualquer texto sem tags na seção de comentários JSDoc será a descrição da função.

No exemplo a seguir, a frase "Uma função que soma dois números" é a descrição da função personalizada com a propriedade id de `ADD`.

```js
/**
 * A function that adds two numbers.
 * @customfunction ADD
 * ...
 */
```

<a id="helpurl"></a>

### <a name="helpurl"></a>@helpurl

Sintaxe: @helpurl _url_

A _url_ fornecida é exibida no Excel.

No exemplo a seguir, o `helpurl` é `www.contoso.com/weatherhelp` .

```js
/**
 * A function which streams the temperature in a town you specify.
 * @customfunction getTemperature
 * @helpurl www.contoso.com/weatherhelp
 * ...
 */
```

<a id="param"></a>

### <a name="param"></a>@param

#### <a name="javascript"></a>JavaScript

Sintaxe de JavaScript: @param {type} nome _descrição_

* `{type}` Especifica as informações de tipo nas chaves. Confira a seção [Tipos](#types) para mais informações sobre os tipos que podem ser usados. Se nenhum tipo for especificado, o tipo padrão `any` será usado.
* `name` Especifica o parâmetro ao qual a marca @param se aplica. É necessário.
* `description` fornece a descrição que aparece no Excel para o parâmetro de função. É opcional.

Para denotar um parâmetro de função personalizado como opcional:

* Coloque colchetes ao redor do nome do parâmetro. Por exemplo: `@param {string} [text] Optional text`.

> [!NOTE]
> O valor padrão para parâmetros opcionais é `null`.

O exemplo a seguir mostra uma função ADD que adiciona dois ou três números, com o terceiro número como um parâmetro opcional.

```js
/**
 * A function which sums two, or optionally three, numbers.
 * @customfunction ADDNUMBERS
 * @param firstNumber {number} First number to add.
 * @param secondNumber {number} Second number to add.
 * @param [thirdNumber] {number} Optional third number you wish to add.
 * ...
 */
```

#### <a name="typescript"></a>TypeScript

Sintaxe de TypeScript: @param nome _descrição_

* `name` Especifica o parâmetro ao qual a marca @param se aplica. É necessário.
* `description` fornece a descrição que aparece no Excel para o parâmetro de função. É opcional.

Confira a seção [Tipos](#types) para mais informações sobre os tipos de parâmetros de função que podem ser usados.

Para denotar um parâmetro de função personalizado como opcional, siga um destes procedimentos:

* Use um parâmetro opcional. Por exemplo: `function f(text?: string)`
* Dê ao parâmetro um valor padrão. Por exemplo: `function f(text: string = "abc")`

Para uma descrição detalhada do @param confira: [JSDoc](https://jsdoc.app/tags-param.html)

> [!NOTE]
> O valor padrão para parâmetros opcionais é `null`.

O exemplo a seguir mostra `add` a função que adiciona dois números.

```ts
/**
 * Adds two numbers.
 * @customfunction 
 * @param first First number
 * @param second Second number
 * @returns The sum of the two numbers.
 */
function add(first: number, second: number): number {
  return first + second;
}
```

<a id="requiresAddress"></a>

### <a name="requiresaddress"></a>@requiresAddress

Indica que o endereço da célula onde a função está sendo avaliada deve ser fornecido.

O último parâmetro da função deve ser do tipo `CustomFunctions.Invocation` ou de um tipo derivado. Quando a função é chamada, a propriedade `address` conterá o endereço.

O exemplo a seguir mostra como usar o `invocation` parâmetro em combinação com `@requiresAddress` para retornar o endereço da célula que chamou sua função personalizada. Consulte [parâmetro de chamada](custom-functions-parameter-options.md#invocation-parameter) para obter mais informações.

```js
/**
 * Return the address of the cell that invoked the custom function. 
 * @customfunction
 * @param {number} first First parameter.
 * @param {number} second Second parameter.
 * @param {CustomFunctions.Invocation} invocation Invocation object. 
 * @requiresAddress 
 */
function getAddress(first, second, invocation) {
  var address = invocation.address;
  return address;
}
```

<a id="returns"></a>

### <a name="returns"></a>@returns

Sintaxe: @returns {_type_}

Fornece o tipo para o valor de retorno.

Se `{type}` for omitido, as informações do tipo TypeScript serão usadas. Se não houver informações de tipo, o tipo será `any`.

O exemplo a seguir mostra a `add` função que usa `@returns` marca.

```ts
/**
 * Adds two numbers.
 * @customfunction 
 * @param first First number
 * @param second Second number
 * @returns The sum of the two numbers.
 */
function add(first: number, second: number): number {
  return first + second;
}
```

<a id="streaming"></a>

### <a name="streaming"></a>@streaming

Usado para indicar que uma função personalizada é uma função de streaming. 

O último parâmetro é do tipo `CustomFunctions.StreamingInvocation<ResultType>` .
A função retorna `void` .

As funções de streaming não retornam valores diretamente, em vez de chamarem `setResult(result: ResultType)` usando o último parâmetro.

Exceções lançadas por uma função de streaming são ignoradas. `setResult()` pode ser chamado com Erro para indicar um resultado de erro. Para obter um exemplo de uma função de streaming e mais informações, confira [, criar uma função de streaming](custom-functions-web-reqs.md#make-a-streaming-function).

As funções de streaming não podem ser marcadas como [@volatile](#volatile).

<a id="volatile"></a>

### <a name="volatile"></a>@volatile

Uma função volátil é aquela cujo resultado não é o mesmo de um momento para o outro, mesmo que não receba argumentos ou os argumentos não mudem. O Excel reavalia células que contenham funções voláteis, juntamente com todos os dependentes, sempre que um cálculo é feito. Por esse motivo, confiar demais em funções voláteis pode retardar o tempo de recálculo; portanto, use-as com moderação.

Funções de streaming não podem ser voláteis.

A função a seguir é volátil e usa `@volatile` a marca.

```js
/**
 * Simulates rolling a 6-sided die.
 * @customfunction
 * @volatile
 */
function roll6sided(): number {
  return Math.floor(Math.random() * 6) + 1;
}
```

---

## <a name="types"></a>Tipos

Especificando um tipo de parâmetro, o Excel converterá valores nesse tipo antes de chamar a função. Se o tipo for `any`, nenhuma conversão será executada.

### <a name="value-types"></a>Tipos de valor

Um valor pode ser representado usando um dos seguintes tipos: `boolean``number``string`.

### <a name="matrix-type"></a>Tipo de matriz

Use um tipo de matriz bidimensional para que o parâmetro ou valor de retorno seja uma matriz de valores. Por exemplo, o tipo `number[][]` indica uma matriz de números. `string[][]` indica uma matriz de cadeias de caracteres.

### <a name="error-type"></a>Tipo de erro

Uma função que não seja de streaming pode indicar um erro retornando um tipo de Erro.

Uma função de streaming pode indicar um erro chamando `setResult()` com um tipo de erro.

### <a name="promise"></a>Promessa

Uma função pode retornar uma promessa, que fornece o valor quando a promessa é resolvida. Se a promessa for rejeitada, será gerado um erro.

### <a name="other-types"></a>Outros tipos

Qualquer outro tipo será tratado como um erro.

## <a name="next-steps"></a>Próximas etapas

Saiba mais sobre [convenções de nomenclatura para funções personalizadas](custom-functions-naming.md). Como alternativa, saiba como [localizar as funções](custom-functions-localize.md) que requerem a [gravação do arquivo JSON à mão](custom-functions-json.md).

## <a name="see-also"></a>Confira também

* [Criar manualmente metadados JSON para funções personalizadas](custom-functions-json.md)
* [Criar funções personalizadas no Excel](custom-functions-overview.md)
