---
title: Diretrizes de nomenização para funções personalizadas Excel
description: Saiba os requisitos para nomes de Excel funções personalizadas e evite armadilhas comuns de nomenbo.
ms.date: 07/08/2021
localization_priority: Normal
ms.openlocfilehash: bfc850fb2a40e7736006930c63489ec7e0c9912b
ms.sourcegitcommit: e570fa8925204c6ca7c8aea59fbf07f73ef1a803
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 08/05/2021
ms.locfileid: "53773346"
---
# <a name="custom-functions-naming-guidelines"></a>Diretrizes de nomenclatura de funções personalizadas

Uma função personalizada é identificada por `id` uma propriedade e no arquivo de `name` metadados JSON.

- A função `id` é usada para identificar exclusivamente funções personalizadas em seu código JavaScript.
- A função `name` é usada como o nome de exibição que aparece para um usuário em Excel.

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

Uma função `name` pode ser diferente da função , como para fins de `id` localização. Em geral, uma função deve permanecer igual à se não houver `name` `id` motivo para diferenciá-las.

Uma função e `name` compartilhar `id` alguns requisitos comuns.

- Uma função só pode usar caracteres A a Z, números de zero a `id` nove, sublinhados e períodos.

- Uma função pode `name` usar quaisquer caracteres alfabéticos Unicode, sublinhados e períodos.

- Ambas funcionam `name` e devem começar com uma letra e ter um limite mínimo de três `id` caracteres.

Excel usa letras maiúsculas para nomes de função integrados (como `SUM` ). Use letras maiúsculas para suas funções personalizadas `name` e como uma prática `id` prática.

Uma função não `name` deve ser igual a:

- Qualquer célula entre A1 e XFD1048576 ou qualquer célula entre R1C1 a R1048576C16384.

- Qualquer Excel função de macro 4.0 (como `RUN` , `ECHO` ).  Para ver uma lista completa dessas funções, consulte [este documento Excel Referência](https://d13ot9o61jdzpp.cloudfront.net/files/Excel%204.0%20Macro%20Functions%20Reference.pdf)de Funções de Macro .

## <a name="naming-conflicts"></a>Conflitos de nomen por nomen

Se sua função for igual a uma função em um complemento que `name` `name` já existe, o **#REF!** error aparecerá em sua workbook.

Para corrigir um conflito de nomeação, altere o `name` no seu complemento e tente a função novamente. Você também pode desinstalar o complemento com o nome conflitante. Ou, se você estiver testando seu complemento em ambientes diferentes, tente usar um namespace diferente para diferenciar sua função (como `NAMESPACE_NAMEOFFUNCTION` ).

## <a name="best-practices"></a>Práticas recomendadas

- Considere adicionar vários argumentos a uma função em vez de criar várias funções com nomes iguais ou semelhantes.
- Evite abreviações ambíguas em nomes de função. A clareza é mais importante do que a brevidade. Escolha um nome `=INCREASETIME` como, em vez de `=INC` .
- Os nomes de função devem indicar a ação da função, como =GETZIPCODE em vez de ZIPCODE.
- Use consistentemente os mesmos verbos para funções que executam ações semelhantes. Por exemplo, use `=DELETEZIPCODE` e , em vez de e `=DELETEADDRESS` `=DELETEZIPCODE` `=REMOVEADDRESS` .
- Ao nomear uma função de streaming, considere adicionar uma nota a esse efeito na descrição da função ou adicionar ao final `STREAM` do nome da função.

[!include[manifest guidance](../includes/manifest-guidance.md)]

## <a name="localizing-function-names"></a>Nomes de função de localização

Você pode localizar seus nomes de função para idiomas diferentes usando arquivos JSON separados e substituir valores no arquivo de manifesto do seu complemento. Evite dar às suas funções uma função interna Excel em outro idioma, pois isso poderia entrar em conflito com funções `id` `name` localizadas.

Para obter informações completas sobre a localização, consulte [Localize custom functions](custom-functions-localize.md)

## <a name="next-steps"></a>Próximas etapas

Saiba mais sobre [as práticas recomendadas de tratamento de erros.](custom-functions-errors.md)

## <a name="see-also"></a>Confira também

* [Criar metadados JSON manualmente para funções personalizadas](custom-functions-json.md)
* [Tutorial de funções personalizadas do Excel](../tutorials/excel-tutorial-create-custom-functions.md)
