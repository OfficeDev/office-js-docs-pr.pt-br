---
ms.date: 04/03/2019
description: Use marcações JSDOC para criar dinamicamente seus metadados JSON de funções personalizadas.
title: Criar metadados JSON para funções personalizadas (visualização)
localization_priority: Priority
ms.openlocfilehash: 2efe2a9a5a83ba60ef327273d5bd599f82916d48
ms.sourcegitcommit: 6d375518c119d09c8d3fb5f0cc4583ba5b20ac03
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 04/18/2019
ms.locfileid: "31914281"
---
# <a name="create-json-metadata-for-custom-functions-preview"></a>Criar metadados JSON para funções personalizadas (visualização)

Quando uma função personalizada do Excel é gravada em JavaScript ou em TypeScript, as marcações JSDoc são usadas para fornecer informações adicionais sobre a função personalizada. As marcações JSDoc são usadas no momento da criação para criar o [arquivo de metadados JSON](custom-functions-json.md). O uso de marcações JSDoc poupa você do esforço de editar manualmente o arquivo de metadados JSON.

Adicione a marcação `@customfunction` nos comentários de código de uma função JavaScript ou TypeScript para marcá-la como uma função personalizada.

Os tipos de parâmetros da função podem ser fornecidos usando a marcação [@param](#param) em JavaScript ou do [Tipo de função](https://www.typescriptlang.org/docs/handbook/functions.html) em TypeScript. Para mais informações, confira a marcação [@param](#param) e a seção [Tipos](#types).

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

Se o último parâmetro da função for do tipo `CustomFunctions.CancelableInvocation`, será considerado `@cancelable` mesmo se a marcação não estiver presente.

Uma função não pode ter ao mesmo tempo as marcações `@cancelable` e `@streaming`.

---
### <a name="customfunction"></a>@customfunction
<a id="customfunction"/>

Sintaxe: @customfunction _id_ _nome_

Especifique esta marcação para tratar a função JavaScript/TypeScript como uma função personalizada do Excel.

Essa marcação é necessária para criar metadados para a função personalizada.

Também deve haver uma chamada para `CustomFunctions.associate("id", functionName);`

#### <a name="id"></a>id 

O id é usado como o identificador invariável da função personalizada armazenada no documento. Ele não deve mudar.

* Se o ID não for fornecido, o nome da função JavaScript/TypeScript será convertido em maiúsculas e os caracteres não permitidos serão removidos.
* O id deve ser exclusivo, para todas as funções personalizadas.
* Os caracteres permitidos são limitados a: A-Z, a-z, 0-9 e ponto (.).

#### <a name="name"></a>nome

Fornece o nome de exibição para a função personalizada. 

* Se nome não for fornecido, o id também será usado como nome.
* Caracteres permitidos: Letras de [caractere Alfabético Unicode](https://www.unicode.org/reports/tr44/tr44-22.html#Alphabetic), números, ponto (.) e sublinhado (\_).
* Deve começar com uma letra.
* O comprimento máximo é de 128 caracteres.

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

#### <a name="typescript"></a>TypeScript

Sintaxe de TypeScript: @param nome _descrição_

* `name` especifica a qual parâmetro a marcação @param se aplica. Obrigatório.
* `description` fornece a descrição que aparece no Excel para o parâmetro de função. Opcional.

Confira [Tipos](##types) para mais informações sobre os tipos de parâmetros de função que podem ser usados.

Para denotar um parâmetro de função personalizado como opcional, siga um destes procedimentos:
* Use um parâmetro opcional. Por exemplo: `function f(text?: string)`
* Dê ao parâmetro um valor padrão. Por exemplo: `function f(text: string = "abc")`

Para uma descrição detalhada do @param confira: [JSDoc](http://usejsdoc.org/tags-param.html)

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

As funções de streaming não retornam valores diretamente, mas em vez disso devem chamar `setResult(result: ResultType)` usando o último parâmetro.

Exceções lançadas por uma função de streaming são ignoradas. `setResult()` pode ser chamado com Erro para indicar um resultado de erro.

Funções de transmissão não podem ser marcadas como [@volatile](#volatile).

---
### <a name="volatile"></a>@volatile
<a id="volatile"/>

Uma função volátil é aquela cujo resultado não pode ser assumido como sendo o mesmo de um momento para o outro, mesmo que não receba argumentos ou que os argumentos não sejam alterados. O Excel reavalia células que contenham funções voláteis, juntamente com todos os dependentes, sempre que um cálculo é feito. Por esse motivo, confiar demais em funções voláteis pode retardar o tempo de recálculo; portanto, use-as com moderação.

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

Uma função de streaming pode indicar um erro chamando setResult () com um tipo de Erro.

### <a name="promise"></a>Promessa

Uma função pode retornar uma Promessa, que fornecerá o valor quando a promessa for resolvida. Se a promessa for rejeitada, então é um erro.

### <a name="other-types"></a>Outros tipos

Qualquer outro tipo será tratado como um erro.

## <a name="see-also"></a>Confira também

* [Metadados de funções personalizadas](custom-functions-json.md)
* [Tempo de execução de funções personalizadas do Excel](custom-functions-runtime.md)
* [Práticas recomendadas de funções personalizadas](custom-functions-best-practices.md).
* [Log de alteração de funções personalizadas](custom-functions-changelog.md)
* [Tutorial de funções personalizadas do Excel](../tutorials/excel-tutorial-create-custom-functions.md)
* [Depuração de funções personalizadas](custom-functions-debugging.md)
