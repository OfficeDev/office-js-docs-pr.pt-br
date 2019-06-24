---
ms.date: 06/18/2019
description: Use tags JSDoc para criar dinamicamente seus metadados JSON de funções personalizadas.
title: Gerar metadados JSON automaticamente para funções personalizadas
localization_priority: Priority
ms.openlocfilehash: a02ca5fd67f29e1997579385e04d045f01e63bdb
ms.sourcegitcommit: 382e2735a1295da914f2bfc38883e518070cec61
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 06/21/2019
ms.locfileid: "35127902"
---
# <a name="autogenerate-json-metadata-for-custom-functions"></a>Gerar metadados JSON automaticamente para funções personalizadas

Quando uma função personalizada do Excel é gravada em JavaScript ou em TypeScript, as marcações JSDoc são usadas para fornecer informações adicionais sobre a função personalizada. As marcações JSDoc são usadas no momento da criação para criar o [arquivo de metadados JSON](custom-functions-json.md). O uso de marcações JSDoc poupa você do esforço de editar manualmente o arquivo de metadados JSON.

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

Adicione a marcação `@customfunction` nos comentários de código de uma função JavaScript ou TypeScript para marcá-la como uma função personalizada.

Os tipos de parâmetros da função podem ser fornecidos usando a marcação [@param](#param) em JavaScript ou do [Tipo de função](https://www.typescriptlang.org/docs/handbook/functions.html) em TypeScript. Para mais informações, confira a marcação [@param](#param) e a seção [Tipos](#types).

### <a name="adding-a-description-to-a-function"></a>Adicionando uma descrição a uma função

A descrição é exibida para o usuário como texto de ajuda quando eles precisam de ajuda para entender o que a função personalizada executa. A descrição não requer nenhuma tag específica. Basta digitar uma breve descrição de texto no comentário JSDoc. Em geral, a descrição é colocada no início da seção de comentários do JSDoc, mas funcionará independentemente de onde seja colocada.

Para ver exemplos das descrições de funções internas, abra o Excel, vá para a guia **Fórmulas** e escolha **Inserir função**. Você pode navegar por todas as descrições de funções e também ver suas próprias funções personalizadas listadas.

No exemplo a seguir, a frase "Calcula o volume de uma esfera." é a descrição da função personalizada.

```JS
/**
/* Calculates the volume of a sphere
/* @customfunction VOLUME
...
 */
```


## <a name="jsdoc-tags"></a>Marcações JSDoc
As seguintes marcações JSDoc possuem suporte em funções personalizadas do Excel:
* [@cancelable](#cancelable)
* [@customfunction](#customfunction) nome de identificação
* [@helpurl](#helpurl) url
* [@param](#param) _{type}_ nome e descrição
* [@requiresAddress](#requiresAddress)
* [@returns](#returns) _{type}_
* [@streaming](#streaming)
* [@volatile](#volatile)

---
### <a name="cancelable"></a>@cancelable
<a id="cancelable"/>

Indica que uma função personalizada deseja executar uma ação quando a função é cancelada.

O último parâmetro da função deve ser do tipo `CustomFunctions.CancelableInvocation`. A função pode atribuir uma função à propriedade `oncanceled` para denotar a ação a ser executada quando a função é cancelada.

Se o último parâmetro da função for do tipo `CustomFunctions.CancelableInvocation`, ela será considerada `@cancelable`, mesmo se a tag não estiver presente.

Uma função não pode ter as tags `@cancelable` e `@streaming` ao mesmo tempo.

---
### <a name="customfunction"></a>@customfunction
<a id="customfunction"/>

Sintaxe: @customfunction _id_ _nome_

Especifique esta marcação para tratar a função JavaScript/TypeScript como uma função personalizada do Excel.

Essa marcação é necessária para criar metadados para a função personalizada.

Também deve haver uma chamada para `CustomFunctions.associate("id", functionName);`

#### <a name="id"></a>id

O `id` é um identificador invariável para a função customizada.

* Se `id` não for fornecido, o nome da função JavaScript/TypeScript será convertido em maiúsculas e os caracteres não permitidos serão removidos.
* O `id` deve ser exclusivo para todas as funções personalizadas.
* Os caracteres permitidos estão limitados a: A-Z, a-z, 0-9, sublinhados (\_) e ponto (.).

#### <a name="name"></a>nome

Fornece a exibição `name` da função personalizada.

* Se o nome não for fornecido, o id também será usado como nome.
* Caracteres permitidos: Letras de [caractere Alfabético Unicode](https://www.unicode.org/reports/tr44/tr44-22.html#Alphabetic), números, ponto (.) e sublinhado (\_).
* Deve começar com uma letra.
* O comprimento máximo é de 128 caracteres.

### <a name="description"></a>description

Uma descrição não exige nenhuma tag específica. Adicione uma descrição a uma função personalizada acrescentando uma frase para descrever o que a função realiza dentro do comentário JSDoc. Por padrão, qualquer texto sem tags na seção de comentários JSDoc será a descrição da função. A descrição aparece para os usuários no Excel quando eles entram na função. No exemplo a seguir, a frase "Uma função que soma dois números" é a descrição da função personalizada com a propriedade id de `SUM`.

```JS
/**
/* @customfunction SUM
/* A function that sums two numbers
...
 */
```

---
### <a name="helpurl"></a>@helpurl
<a id="helpurl"/>

Sintaxe: @helpurl _url_

A _url_ fornecida é exibida no Excel.

---
### <a name="param"></a>@param
<a id="param"/>

#### <a name="javascript"></a>JavaScript

Sintaxe de JavaScript: @param {type} nome _descrição_

* `{type}` deve especificar a informação de tipo entre chaves. Confira [Tipos](##types) para mais informações sobre os tipos que podem ser usados. Opcional: se não especificado, o tipo `any` será usado.
* `name` especifica a qual parâmetro a marcação @param se aplica. Obrigatório.
* `description` fornece a descrição que aparece no Excel para o parâmetro de função. Opcional.

Para denotar um parâmetro de função personalizado como opcional:
* Coloque colchetes ao redor do nome do parâmetro. Por exemplo: `@param {string} [text] Optional text`.

> [!NOTE]
> O valor padrão para parâmetros opcionais é `null`.

#### <a name="typescript"></a>TypeScript

Sintaxe de TypeScript: @param nome _descrição_

* `name` especifica a qual parâmetro a marcação @param se aplica. Obrigatório.
* `description` fornece a descrição que aparece no Excel para o parâmetro de função. Opcional.

Confira [Tipos](##types) para mais informações sobre os tipos de parâmetros de função que podem ser usados.

Para denotar um parâmetro de função personalizado como opcional, siga um destes procedimentos:
* Use um parâmetro opcional. Por exemplo: `function f(text?: string)`
* Dê ao parâmetro um valor padrão. Por exemplo: `function f(text: string = "abc")`

Para uma descrição detalhada do @param confira: [JSDoc](https://jsdoc.app/tags-param.html)

> [!NOTE]
> O valor padrão para parâmetros opcionais é `null`.

---
### <a name="requiresaddress"></a>@requiresAddress
<a id="requiresAddress"/>

Indica que o endereço da célula onde a função está sendo avaliada deve ser fornecido. 

O último parâmetro da função deve ser do tipo `CustomFunctions.Invocation` ou de um tipo derivado. Quando a função é chamada, a propriedade `address` conterá o endereço.

---
### <a name="returns"></a>@returns
<a id="returns"/>

Sintaxe: @returns {_type_}

Fornece o tipo para o valor de retorno.

Se `{type}` for omitido, as informações do tipo TypeScript serão usadas. Se não houver informações de tipo, o tipo será `any`.

---
### <a name="streaming"></a>@streaming
<a id="streaming"/>

Usado para indicar que uma função personalizada é uma função de streaming. 

O último parâmetro deve ser do tipo `CustomFunctions.StreamingInvocation<ResultType>`.
A função deve retornar `void`.

As funções de streaming não retornam valores diretamente, mas devem chamar `setResult(result: ResultType)` usando o último parâmetro.

Exceções lançadas por uma função de streaming são ignoradas. `setResult()` pode ser chamado com Erro para indicar um resultado de erro.

As funções de streaming não podem ser marcadas como [@volatile](#volatile).

---
### <a name="volatile"></a>@volatile
<a id="volatile"/>

Uma função volátil é aquela cujo resultado não é o mesmo de um momento para o outro, mesmo que não receba argumentos ou os argumentos não mudem. O Excel reavalia células que contenham funções voláteis, juntamente com todos os dependentes, sempre que um cálculo é feito. Por esse motivo, confiar demais em funções voláteis pode retardar o tempo de recálculo; portanto, use-as com moderação.

Funções de streaming não podem ser voláteis.

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

Uma função pode retornar uma Promessa, que fornecerá o valor quando a promessa for resolvida. Se a promessa for rejeitada, então é um erro.

### <a name="other-types"></a>Outros tipos

Qualquer outro tipo será tratado como um erro.

## <a name="next-steps"></a>Próximas etapas
Saiba mais sobre [convenções de nomenclatura para funções personalizadas](custom-functions-naming.md). Como alternativa, saiba como [localizar as funções](custom-functions-localize.md) que requerem a [gravação do arquivo JSON à mão](custom-functions-json.md).

## <a name="see-also"></a>Confira também

* [Metadados de funções personalizadas](custom-functions-json.md)
* [Práticas recomendadas de funções personalizadas](custom-functions-best-practices.md).
* [Criar funções personalizadas no Excel](custom-functions-overview.md)
