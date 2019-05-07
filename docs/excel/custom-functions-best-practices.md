---
ms.date: 05/06/2019
description: Saiba mais sobre as práticas recomendadas para o desenvolvimento de funções personalizadas para Excel.
title: Práticas recomendadas para funções personalizadas
localization_priority: Normal
ms.openlocfilehash: 7369faa463966dd309258bf431eae8719407be38
ms.sourcegitcommit: ff73cc04e5718765fcbe74181505a974db69c3f5
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/06/2019
ms.locfileid: "33628132"
---
# <a name="custom-functions-best-practices"></a>Práticas recomendadas para funções personalizadas

Este artigo descreve as práticas recomendadas para o desenvolvimento de funções personalizadas para Excel.

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

## <a name="associating-function-names-with-json-metadata"></a>Associar os nomes de função com metadados JSON

Conforme descrito no artigo [visão geral de funções personalizados](custom-functions-overview.md), um projeto de funções personalizados deve incluir um arquivo JSON de metadados e um arquivo de script (JavaScript ou TypeScript) para formar uma função completa. Se você estiver usando `yo office` os metadados JSON podem ser gerados a partir dos comentários de código. Caso contrário, você precisará criar o arquivo de metadados JSON manualmente.

Para que uma função funcione corretamente, você precisa associar a propriedade da `id` função à implementação do JavaScript. Verifique se há uma associação, caso contrário, a função não será chamada. O exemplo de código a seguir mostra como fazer a Associação usando `CustomFunctions.associate()` o método. A amostra define a função personalizada `add` e associa com o objeto no arquivo de metadados JSON onde o valor da `id` propriedade é **adicionar**.

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

O JSON a seguir mostra os metadados JSON que estão associados ao código JavaScript da função personalizada anterior.

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
    },
  ]
}
```


Lembre-se das seguintes práticas recomendadas quando criar funções personalizadas no arquivo JavaScript e especificar as informações correspondentes no arquivo de metadados JSON.

* No arquivo de metadados JSON, verifique se o valor de cada propriedade `id` contém apenas caracteres alfanuméricos e pontos.

* No arquivo de metadados JSON, garanta que o valor de cada propriedade `id` seja exclusivo dentro do escopo do arquivo. Ou seja, nenhum objeto de duas funções no arquivo de metadados deve ter o mesmo valor `id`.

* Não altere o valor de uma propriedade `id` no arquivo de metadados JSON, depois de mapeá-lo para um nome de função JavaScript correspondente. Para alterar o nome da função que os usuários finais visualizam no Excel, atualize a propriedade `name` no arquivo de metadados JSON. No entanto, nunca altere o valor de uma propriedade `id` depois de estabelecida.

* No arquivo JavaScript, especifique uma associação de função personalizada usando `CustomFunctions.associate` após cada função.

O exemplo a seguir mostra os metadados JSON que correspondem às funções definidas nesse exemplo de código JavaScript. Os `id` valores `name` de propriedade e estão em letras maiúsculas, o que é uma prática recomendada ao descrever suas funções personalizadas. Você só precisará adicionar esse JSON se estiver preparando seu próprio arquivo JSON manualmente e não usando a autogeração. Para obter mais informações sobre a autogeração, consulte [criar metadados JSON para funções personalizadas](custom-functions-json-autogeneration.md).

```json
{
  "$schema": "https://developer.microsoft.com/en-us/json-schemas/office-js/custom-functions.schema.json",
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

## <a name="additional-considerations"></a>Considerações adicionais

Evite acessar o modelo de objeto de documento (DOM) direta ou indiretamente (por exemplo, usando jQuery) de sua função personalizada. No Excel para Windows, onde as funções personalizadas usam o [tempo de execução do JavaScript](custom-functions-runtime.md), as funções personalizadas não podem acessar o DOM.

## <a name="next-steps"></a>Próximas etapas
Saiba como [realizar solicitações da Web com funções personalizadas](custom-functions-web-reqs.md).

## <a name="see-also"></a>Confira também

* [Gerar automaticamente metadados JSON para funções personalizadas](custom-functions-json-autogeneration.md)
* [Metadados de funções personalizadas](custom-functions-json.md)
* [Criar funções personalizadas no Excel](custom-functions-overview.md)
