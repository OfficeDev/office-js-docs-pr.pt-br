---
ms.date: 07/09/2019
description: Autentique usuários usando funções personalizadas no Excel.
title: Autenticação para funções personalizadas
localization_priority: Priority
ms.openlocfilehash: 74e1524eaf9c5328754fee8c225cd5aca83188da
ms.sourcegitcommit: bb44c9694f88cde32ffbb642689130db44456964
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 07/17/2019
ms.locfileid: "35771467"
---
# <a name="authentication-for-custom-functions"></a>Autenticação para funções personalizadas

Em alguns cenários, sua função personalizada precisará autenticar o usuário para acessar recursos protegidos. Embora as funções personalizadas não exijam um método específico de autenticação, você deve estar ciente de que as funções personalizadas são executadas em um tempo de execução separado no painel de tarefas e em outros elementos da interface do usuário do seu suplemento. Por causa disso, você precisará passar dados alternando entre os dois tempos de execução usando o objeto `OfficeRuntime.storage` e a API de Caixa de Diálogo.

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

## <a name="officeruntimestorage-object"></a>Objeto OfficeRuntime.storage

O tempo de execução de funções personalizadas não tem um objeto `localStorage` disponível na janela global, onde você normalmente pode armazenar dados. Em vez disso, você deve compartilhar dados entre funções personalizadas e painéis de tarefas usando o [OfficeRuntime.storage](/javascript/api/office-runtime/officeruntime.storage) para definir e obter dados.

Além disso, há um benefício em usar o objeto `storage`; Ele usa um ambiente de sandbox seguro para que seus dados não possam ser acessados por outros suplementos.

### <a name="suggested-usage"></a>Uso sugerido

Quando você precisar autenticar a partir do painel de tarefas ou de uma função personalizada, verifique `storage` para ver se o token de acesso já foi adquirido. Caso contrário, use a API de caixa de diálogo para autenticar o usuário, recuperar o token de acesso e, em seguida, armazenar o token em `storage` para uso futuro.

## <a name="dialog-api"></a>API de Caixa de Diálogo

Se um token não existir, você deverá usar a API de diálogo para solicitar que o usuário faça logon. Depois que um usuário insere suas credenciais, o token de acesso resultante pode ser armazenado em `storage`.

> [!NOTE]
> O tempo de execução de funções personalizadas usa um objeto Dialog que é um pouco diferente do objeto Dialog no tempo de execução do mecanismo do navegador usado pelos painéis de tarefas. Ambos são chamados de "API de Caixa de Diálogo", mas usam `OfficeRuntime.Dialog` para autenticar usuários no tempo de execução de funções personalizadas.

Para obter informações sobre como usar o objeto `Dialog`, consulte a [Caixa de Diálogo Funções Personalizadas](/office/dev/add-ins/excel/custom-functions-dialog).

Ao visualizar o processo de autenticação como um todo, pode ser útil pensar no painel de tarefas e nos elementos de IU de seu suplemento, bem como pensar nas funções personalizadas de seu complemento como entidades separadas que podem se comunicar entre si por meio de `OfficeRuntime.storage`.

O diagrama a seguir descreve esse processo básico. Observe que a linha pontilhada indica que, embora executem ações separadas, as funções personalizadas e o painel de tarefas do seu suplemento fazem parte do seu suplemento como um todo.

1. As chamadas de função personalizada de uma célula são emitdas por você em uma pasta de trabalho do Excel.
2. A função personalizada usa `Dialog` para passar suas credenciais de usuário para um site.
3. Esse site, em seguida, retorna um token de acesso para a função personalizada.
4. Sua função personalizada, em seguida, define esse token de acesso para `storage`.
5. O painel de tarefas do seu suplemento acessa o token a partir de `storage`.

![Diagrama da função personalizada usando a API de caixa de diálogo para obter o token de acesso e, em seguida, compartilhar o token com o painel de tarefas por meio da API OfficeRuntime.storage. ](../images/authentication-diagram.png "Diagrama de autenticação.")

## <a name="storing-the-token"></a>Armazenando o token

Os exemplos a seguir são do exemplo de código [Usando OfficeRuntime.storage em funções personalizadas](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Excel-custom-functions/AsyncStorage). Consulte este exemplo de código para obter um exemplo completo de compartilhamento de dados entre funções personalizadas e o painel de tarefas.

Se a função personalizada for autenticada, ela receberá o token de acesso e precisará armazená-lo em `storage`. O exemplo de código a seguir mostra como chamar o método `storage.setItem` para armazenar um valor. A função `storeValue` é uma função personalizada que, para fins de exemplo, armazena um valor do usuário. Você pode modificá-la para que seja armazenado qualquer valor de token que você precise.

```js
/**
 * Stores a key-value pair into OfficeRuntime.storage.
 * @customfunction
 * @param {string} key Key of item to put into storage.
 * @param {*} value Value of item to put into storage.
 */
function storeValue(key, value) {
  return OfficeRuntime.storage.setItem(key, value).then(function (result) {
      return "Success: Item with key '" + key + "' saved to storage.";
  }, function (error) {
      return "Error: Unable to save item with key '" + key + "' to storage. " + error;
  });
}
```

Quando o painel de tarefas precisa do token de acesso, ele pode recuperar o token de `storage`. O exemplo de código a seguir mostra como usar o método `storage.getItem` para recuperar o token.

```js
/**
 * Read a token from storage.
 * @customfunction GETTOKEN
 */
function receiveTokenFromCustomFunction() {
  var key = "token";
  var tokenSendStatus = document.getElementById('tokenSendStatus');
  OfficeRuntime.storage.getItem(key).then(function (result) {
     tokenSendStatus.value = "Success: Item with key '" + key + "' read from storage.";
     document.getElementById('tokenTextBox2').value = result;
  }, function (error) {
     tokenSendStatus.value = "Error: Unable to read item with key '" + key + "' from storage. " + error;
  });
}
```

## <a name="general-guidance"></a>Orientação geral

Os Suplementos do Office são baseados na Web e você pode usar qualquer técnica de autenticação da Web. Não há um padrão ou método específico que você deva seguir para implementar sua própria autenticação com funções personalizadas. Você pode querer consultar a documentação sobre vários padrões de autenticação, começando com [este artigo sobre a autorização por serviços externos](/office/dev/add-ins/develop/auth-external-add-ins?view=office-js).  

Evite usar os seguintes locais para armazenar dados ao desenvolver funções personalizadas:  

- `localStorage`: Funções personalizadas não têm acesso ao objeto global `window` e, portanto, não têm acesso aos dados armazenados em `localStorage`.
- `Office.context.document.settings`: Esse local não é seguro, e informações podem ser extraídas por qualquer pessoa usando o suplemento.

## <a name="next-steps"></a>Próximas etapas
Aprenda sobre a [API de caixa de diálogo para funções personalizadas](custom-functions-dialog.md).

## <a name="see-also"></a>Confira também

* [Arquitetura de funções personalizadas](custom-functions-architecture.md)
* [Receber e tratar dados com funções personalizadas](custom-functions-web-reqs.md)
* [Tempo de execução de funções personalizadas do Excel](custom-functions-runtime.md)
* [Tutorial de funções personalizadas do Excel](excel-tutorial-custom-functions.md)
