---
title: Diretrizes de nomenização para funções personalizadas no Excel
description: Saiba os requisitos para nomes de Excel funções personalizadas e evite armadilhas comuns de nomenbo.
ms.date: 07/08/2021
ms.localizationpriority: medium
ms.openlocfilehash: 629ed7000046a2cf543e0ac9e398c349666a67c1
ms.sourcegitcommit: 968d637defe816449a797aefd930872229214898
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/23/2022
ms.locfileid: "63744515"
---
# <a name="custom-functions-naming-guidelines"></a>Diretrizes de nomenclatura de funções personalizadas

Uma função personalizada é identificada por uma propriedade `id` e `name` no arquivo de metadados JSON.

- A função `id` é usada para identificar exclusivamente funções personalizadas em seu código JavaScript.
- A função `name` é usada como o nome de exibição que aparece para um usuário no Excel.

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

Uma função `name` pode ser diferente da função `id`, como para fins de localização. Em geral, uma função `name` deve permanecer igual à `id` se não houver motivo para diferenciá-las.

Uma função e compartilhar `name` `id` alguns requisitos comuns.

- Uma função só pode `id` usar caracteres A a Z, números de zero a nove, sublinhados e períodos.

- Uma função pode `name` usar quaisquer caracteres alfabéticos Unicode, sublinhados e períodos.

- Ambas funcionam `name` e `id` devem começar com uma letra e ter um limite mínimo de três caracteres.

Excel usa letras maiúsculas para nomes de função integrados (como `SUM`). Use letras maiúsculas para suas funções personalizadas `name` e `id` como uma prática prática.

Uma função não `name` deve ser igual a:

- Qualquer célula entre A1 e XFD1048576 ou qualquer célula entre R1C1 a R1048576C16384.

- Qualquer Excel função de macro 4.0 (como `RUN`, `ECHO`).  Para ver uma lista completa dessas funções, consulte [este documento de referência de funções de macro Excel macro](https://d13ot9o61jdzpp.cloudfront.net/files/Excel%204.0%20Macro%20Functions%20Reference.pdf).

## <a name="naming-conflicts"></a>Conflitos de nomen por nomen

Se sua função `name` for igual a uma função `name` em um complemento que já existe, o #REF **!** error aparecerá em sua workbook.

Para corrigir um conflito de nomeação, altere `name` o no seu complemento e tente a função novamente. Você também pode desinstalar o complemento com o nome conflitante. Ou, se você estiver testando seu complemento em ambientes diferentes, tente usar um namespace diferente para diferenciar sua função (como `NAMESPACE_NAMEOFFUNCTION`).

## <a name="best-practices"></a>Práticas recomendadas

- Considere adicionar vários argumentos a uma função em vez de criar várias funções com nomes iguais ou semelhantes.
- Evite abreviações ambíguas em nomes de função. A clareza é mais importante do que a brevidade. Escolha um nome como, em `=INCREASETIME` vez de `=INC`.
- Os nomes de função devem indicar a ação da função, como =GETZIPCODE em vez de ZIPCODE.
- Use consistentemente os mesmos verbos para funções que executam ações semelhantes. Por exemplo, use `=DELETEZIPCODE` e `=DELETEADDRESS`, em vez de `=DELETEZIPCODE` e `=REMOVEADDRESS`.
- Ao nomear uma função de streaming, considere adicionar uma nota a `STREAM` esse efeito na descrição da função ou adicionar ao final do nome da função.

[!include[manifest guidance](../includes/manifest-guidance.md)]

## <a name="localizing-function-names"></a>Nomes de função de localização

Você pode localizar seus nomes de função para idiomas diferentes usando arquivos JSON separados e substituir valores no arquivo de manifesto do seu complemento. Evite dar às suas funções uma `id` `name` ou que seja uma função interna Excel em outro idioma, pois isso poderia entrar em conflito com funções localizadas.

Para obter informações completas sobre a localização, consulte [Localize custom functions](custom-functions-localize.md)

## <a name="next-steps"></a>Próximas etapas

Saiba mais sobre [as práticas recomendadas de tratamento de erros](custom-functions-errors.md).

## <a name="see-also"></a>Confira também

* [Criar metadados JSON manualmente para funções personalizadas](custom-functions-json.md)
* [Tutorial de funções personalizadas do Excel](../tutorials/excel-tutorial-create-custom-functions.md)
