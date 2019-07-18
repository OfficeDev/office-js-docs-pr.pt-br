---
ms.date: 07/10/2019
description: Saiba mais sobre os nomes de funções personalizadas do Excel e evite armadilhas comuns de nomeação.
title: Diretrizes de nomenclatura para funções personalizadas no Excel
localization_priority: Normal
ms.openlocfilehash: 79d0bfb069fe5abefeb6d0e88428d0728f3869e3
ms.sourcegitcommit: bb44c9694f88cde32ffbb642689130db44456964
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 07/17/2019
ms.locfileid: "35771530"
---
# <a name="naming-guidelines"></a>Diretrizes de nomenclatura

Uma função personalizada é identificada por uma propriedade **ID** e **nome** no arquivo de metadados JSON.

- A função `id` é usada para identificar exclusivamente as funções personalizadas no seu código JavaScript.
- A função `name` é usada como o nome de exibição que aparece para um usuário no Excel.

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

Uma função `name` pode ser diferente da função `id`, como para fins de localização. Em geral, uma função `name` deve permanecer igual a `id` se não houver um motivo convincente para elas diferirem.

Uma função `name` e `id` compartilhar alguns requisitos comuns:

- Uma função `id` pode usar apenas caracteres de a a Z, números de zero a nove, sublinhados e pontos.

- Uma função `name` pode usar caracteres alfabéticos Unicode, sublinhados e pontos.

- Ambas funcionam `name` e `id` devem começar com uma letra e ter um limite mínimo de três caracteres.

O `SUM`Excel usa letras maiúsculas para nomes de função internos (como). Portanto, considere o uso de letras maiúsculas para a `name` função `id` personalizada e como uma prática recomendada.

Uma função não `name` deve ser nomeada da mesma forma:

- Qualquer célula entre a1 e XFD1048576 ou qualquer célula entre L1C1 e R1048576C16384.

- Qualquer função de macro do Excel 4,0 ( `RUN`como `ECHO`,).  Para obter uma lista completa dessas funções, consulte [este documento de referência de funções de macro do Excel](https://d13ot9o61jdzpp.cloudfront.net/files/Excel%204.0%20Macro%20Functions%20Reference.pdf).

## <a name="naming-conflicts"></a>Conflitos de nomenclatura

Se sua função `name` for igual a uma função `name` em um suplemento que já existe, o **#REF!** o erro aparecerá na sua pasta de trabalho.

Para corrigir um conflito de nomenclatura, altere `name` o em seu suplemento e repita a função. Você também pode desinstalar o suplemento com o nome conflitante. Ou, se você estiver testando seu suplemento em diferentes ambientes, tente usar um namespace diferente para diferenciar sua função (como `NAMESPACE_NAMEOFFUNCTION`).

## <a name="best-practices"></a>Práticas recomendadas

- Considere adicionar vários argumentos a uma função em vez de criar várias funções com nomes iguais ou semelhantes.
- Os nomes de função devem indicar a ação da função, como `=GETZIPCODE` em vez `ZIPCODE`de.
- Evite abreviações ambíguas em nomes de funções. A clareza é mais importante do que a brevidade. Escolha um nome como `=INCREASETIME` em vez `=INC`de.
- Use consistentemente os mesmos verbos para funções que executam ações semelhantes. Por exemplo, use `=DELETEZIPCODE` e `=DELETEADDRESS`, em vez `=DELETEZIPCODE` de `=REMOVEADDRESS`e.
- Ao nomear uma função de streaming, considere adicionar uma nota a esse efeito na descrição da função ou adicionar `STREAM` ao final do nome da função.

## <a name="localizing-function-names"></a>Localizando nomes de função

Você pode localizar seus nomes de função para idiomas diferentes usando arquivos JSON separados e substituir valores no arquivo de manifesto do seu suplemento. Como prática recomendada, evite dar às funções uma `id` ou `name` que é uma função interna do Excel em outro idioma, pois isso pode causar conflito com funções localizadas.

Para obter informações completas sobre a localização, consulte [localizar funções personalizadas](custom-functions-localize.md)

## <a name="next-steps"></a>Próximas etapas
Saiba mais sobre [as práticas recomendadas de tratamento de erros](custom-functions-errors.md).

## <a name="see-also"></a>Confira também

* [Metadados de funções personalizadas](custom-functions-json.md)
* [Tutorial de funções personalizadas do Excel](../tutorials/excel-tutorial-create-custom-functions.md)
* [Tempo de execução de funções personalizadas do Excel](custom-functions-runtime.md)
