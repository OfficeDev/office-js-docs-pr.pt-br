---
ms.date: 02/08/2019
description: Saiba mais sobre os nomes de funções personalizadas do Excel e evite armadilhas comuns de nomeação.
title: Diretrizes de nomenclatura para funções personalizadas no Excel (visualização)
localization_priority: Normal
ms.openlocfilehash: 954753c35d2df59093661e3b8e92adfa1302e595
ms.sourcegitcommit: 8e7b7b0cfb68b91a3a95585d094cf5f5ffd00178
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/09/2019
ms.locfileid: "30512836"
---
# <a name="naming-guidelines"></a>Diretrizes de nomenclatura

Uma função personalizada é identificada por uma propriedade **ID** e **nome** no arquivo de metadados JSON. A ID da função é usada para identificar exclusivamente as funções personalizadas no seu código JavaScript. O nome da função é usado como o nome de exibição que aparece para um usuário no Excel. Um nome de função pode ser diferente da ID da função, como para fins de localização. Mas em geral, ela deve permanecer igual à ID se não houver uma razão convincente para elas diferirem.

Os nomes de função e as IDs de função compartilham alguns requisitos comuns:

- As IDs de função só podem usar caracteres de A a Z, números de zero a nove, sublinhados e pontos.

- Os nomes de função podem usar caracteres alfabéticos Unicode, sublinhados e pontos.

- Eles devem começar com uma letra e ter um limite mínimo de três caracteres.

O `SUM`Excel usa letras maiúsculas para nomes de função internos (como). Portanto, considere o uso de letras maiúsculas para seus nomes de função personalizada e IDs de função como uma prática recomendada.

Os nomes de função não devem ser nomeados da mesma forma:

- Qualquer célula entre a1 e XFD1048576 ou qualquer célula entre L1C1 e R1048576C16384.

- Qualquer função de macro do Excel 4,0 ( `RUN`como `ECHO`,).  Para obter uma lista completa dessas funções, consulte [Este artigo](https://www.microsoft.com/en-us/download/details.aspx?id=1465).

## <a name="naming-conflicts"></a>Conflitos de nomenclatura

Se o nome da função for igual ao nome de uma função em um suplemento que já existe, o **#REF!** o erro aparecerá na sua pasta de trabalho.

Para corrigir um conflito de nomes, altere o nome no suplemento e repita a função. Você também pode desinstalar o suplemento com o nome conflitante. Ou, se você estiver testando seu suplemento em diferentes ambientes, tente usar um namespace diferente para diferenciar sua função (como NAMESPACE_NAMEOFFUNCTION).

Considere também como você gostaria que as pessoas usem as funções dentro do seu suplemento. Em muitos casos, faz sentido adicionar vários argumentos a uma função, em vez de criar várias funções com nomes iguais ou semelhantes.

## <a name="see-also"></a>Confira também

* [Metadados de funções personalizadas](custom-functions-json.md)
* [Práticas recomendadas de funções personalizadas](custom-functions-best-practices.md).
* [Tutorial de funções personalizadas do Excel](../tutorials/excel-tutorial-create-custom-functions.md)
* [Tempo de execução de funções personalizadas do Excel](custom-functions-runtime.md)
