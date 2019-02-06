---
ms.date: 1/29/2019
description: Autentica usuários usando as funções personalizadas no Excel.
title: Autenticação para funções personalizadas
ms.openlocfilehash: 0e42dbc93cb545660a8dbaae5bdb48724f3b7376
ms.sourcegitcommit: 33dcf099c6b3d249811580d67ee9b790c0fdccfb
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 02/05/2019
ms.locfileid: "29745398"
---
# <a name="authentication"></a>Autenticação

Em alguns cenários que sua função personalizada será necessário para autenticar o usuário para acessar recursos protegidos. Enquanto as funções personalizadas não exige um método específico de autenticação, você deve estar ciente de que funções personalizadas é executado em um tempo de execução separado do painel de tarefas e outros elementos de interface do usuário do seu suplemento. Dessa forma, você precisará passar dados bidirecionalmente entre os dois tempos de execução usando o `AsyncStorage` objeto e a API de diálogo.
  
## <a name="asyncstorage-object"></a>Objeto AsyncStorage

O tempo de execução de funções personalizadas não tem um `localStorage` objeto disponível na janela global, onde você pode armazenar dados normalmente. Em vez disso, você deve compartilhar dados entre funções personalizadas e painéis de tarefas, usando [OfficeRuntime.AsyncStorage](https://docs.microsoft.com/javascript/api/office-runtime/officeruntime.asyncstorage) para definir e obter dados. 

Além disso, há um benefício usando `AsyncStorage`; Ele usa um ambiente seguro segura para que seus dados não podem ser acessados por outros complementos.  

### <a name="suggested-usage"></a>Uso sugerido

Quando você precisa para autenticar a partir do painel de tarefas ou uma função personalizada, verifique AsyncStorage para ver se o token de acesso já foi adquirido. Caso contrário, use a caixa de diálogo API para autenticar o usuário, recupere o token de acesso e armazene o token no AsyncStorage para uso futuro.

## <a name="dialog-api"></a>API de diálogo

Se não existir um token, você deve usar a API de diálogo pedir ao usuário para entrar. Depois que um usuário digita suas credenciais, o token de acesso resultantes pode ser armazenado em `AsyncStorage`.

> [!NOTE]
> O tempo de execução de funções personalizadas usa um objeto de diálogo que é ligeiramente diferente do objeto Dialog no runtime usado pelo painéis de tarefas. Estiver ambos conhecidos como "Diálogo API", mas use `Officeruntime.Dialog` para autenticar usuários em tempo de execução de funções personalizadas.

Para obter informações sobre como usar o `OfficeRuntime.Dialog`, consulte o [tempo de execução de funções personalizadas](https://docs.microsoft.com/en-us/office/dev/add-ins/excel/custom-functions-runtime?view=office-js#displaying-a-dialog-box).

Quando envisioning o processo de autenticação inteira como um todo, talvez seja útil pensar o painel de tarefas e elementos de interface do usuário do seu suplemento e custom funciona partes de seu suplemento como entidades separadas que podem se comunicar entre si por meio de `AsyncStorage`.

O diagrama a seguir descreve esse processo básico. Observe que a linha pontilhada indica que enquanto executam ações separadas, funções personalizadas e painel de tarefas do seu suplemento fazem parte do seu suplemento como um todo.

1. Você emite uma chamada de função personalizada de uma célula em uma planilha do Excel.
2. A função personalizada usa `Officeruntime.Dialog` passar suas credenciais de usuário para um site.
3. Este site, em seguida, retorna um token de acesso para a função personalizada.
4. Sua função personalizada define este token de acesso para o `AsyncStorage`.
5. Painel de tarefas do suplemento acessa o token de `AsyncStorage`.

![Diagrama de funções personalizadas, OfficeRuntime e painéis de tarefas trabalhando juntos.] (../images/Authdiagram.png "Diagrama de autenticação.")

## <a name="general-guidance"></a>Diretrizes gerais

Suplementos do Office são baseados na web e você pode usar qualquer técnica de autenticação da web. Não há nenhum padrão específico ou método que você deve seguir para implementar sua própria autenticação com funções personalizadas. Você poderá consultar a documentação sobre vários padrões de autenticação, começando com [Este artigo sobre como autorizar via serviços externos](https://docs.microsoft.com/en-us/office/dev/add-ins/develop/auth-external-add-ins?view=office-js).  

Evite o uso dos seguintes locais para armazenar dados ao desenvolver funções personalizadas:  

- `localStorage`: Funções personalizadas não têm acesso ao global `window` objeto e, portanto, não têm acesso aos dados armazenados em `localStorage`.
- `Office.context.document.settings`: Este local não é seguro e informações podem ser extraídas pela pessoa que usar o suplemento.

## <a name="see-also"></a>Confira também

* [Metadados de funções personalizadas](custom-functions-json.md)
* [Tempo de execução de funções personalizadas do Excel](custom-functions-runtime.md)
* [Práticas recomendadas de funções personalizadas](custom-functions-best-practices.md).
* [Tutorial de funções personalizadas do Excel](excel-tutorial-custom-functions.md)
