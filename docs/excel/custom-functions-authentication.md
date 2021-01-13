---
ms.date: 05/17/2020
description: Autenticar usuários usando funções personalizadas no Excel que não usam o painel de tarefas.
title: Autenticação para funções personalizadas sem interface do usuário
localization_priority: Normal
ms.openlocfilehash: bca3cd422330b6499e18c31ef8d7da6def81b546
ms.sourcegitcommit: d28392721958555d6edea48cea000470bd27fcf7
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 01/13/2021
ms.locfileid: "49839856"
---
# <a name="authentication-for-ui-less-custom-functions"></a>Autenticação para funções personalizadas sem interface do usuário

Em alguns cenários, sua função personalizada que não usa um painel de tarefas ou outros elementos da interface do usuário (função personalizada sem interface do usuário) precisará autenticar o usuário para acessar recursos protegidos. Esteja ciente de que funções personalizadas sem interface do usuário são executados em um tempo de execução somente JavaScript. Por isso, você precisará passar dados entre o tempo de execução somente JavaScript e o tempo de execução típico do mecanismo do navegador usado pela maioria dos complementos usando o objeto e a API de Caixa de `OfficeRuntime.storage` Diálogo.

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

[!include[Shared runtime note](../includes/shared-runtime-note.md)]

## <a name="officeruntimestorage-object"></a>Objeto OfficeRuntime.storage

O tempo de execução somente JavaScript usado por funções personalizadas sem interface do usuário não tem um objeto disponível na janela global, onde você normalmente `localStorage` armazena dados. Em vez disso, você deve compartilhar dados entre funções personalizadas sem interface do usuário e painéis de tarefas usando [OfficeRuntime.storage](/javascript/api/office-runtime/officeruntime.storage) para definir e obter dados.

### <a name="suggested-usage"></a>Uso sugerido

Quando você precisar se autenticar de uma função personalizada sem interface do usuário, verifique se o `storage` token de acesso já foi adquirido. Caso contrário, use a API de caixa de diálogo para autenticar o usuário, recuperar o token de acesso e, em seguida, armazenar o token em `storage` para uso futuro.

## <a name="dialog-api"></a>API de Caixa de Diálogo

Se um token não existir, você deverá usar a API de diálogo para solicitar que o usuário faça logon. Depois que um usuário insere suas credenciais, o token de acesso resultante pode ser armazenado em `storage`.

> [!NOTE]
> O tempo de execução somente JavaScript usa um objeto Dialog ligeiramente diferente do objeto Dialog no tempo de execução do mecanismo do navegador usado pelos painéis de tarefas. Ambos são chamados de "API da Caixa de Diálogo", mas são usadas para autenticar usuários no tempo de execução `OfficeRuntime.Dialog` somente JavaScript.

O diagrama a seguir descreve esse processo básico. A linha pontilhada indica que funções personalizadas sem interface do usuário e o painel de tarefas do seu complemento fazem parte do seu complemento como um todo, embora usem tempos de execução separados.

1. Você emmitiu uma chamada de função personalizada sem interface do usuário de uma célula em uma planilha do Excel.
2. A função personalizada sem interface do usuário usa `Dialog` para passar suas credenciais de usuário para um site.
3. Em seguida, este site retorna um token de acesso à função personalizada sem interface do usuário.
4. Sua função personalizada sem interface do usuário define esse token de acesso como `storage` .
5. O painel de tarefas do seu suplemento acessa o token a partir de `storage`.

![Diagrama de função personalizada usando a API de caixa de diálogo para obter o token de acesso e compartilhar o token com o painel de tarefas por meio da API OfficeRuntime.storage.](../images/authentication-diagram.png "Diagrama de autenticação.")

## <a name="storing-the-token"></a>Armazenando o token

Os exemplos a seguir são do exemplo de código [Usando OfficeRuntime.storage em funções personalizadas](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Excel-custom-functions/AsyncStorage). Consulte este exemplo de código para ver um exemplo completo de compartilhamento de dados entre funções personalizadas sem interface do usuário e o painel de tarefas.

Se a função personalizada sem interface do usuário autenticar, ela receberá o token de acesso e precisará armazená-lo. `storage` O exemplo de código a seguir mostra como chamar o método `storage.setItem` para armazenar um valor. A função é uma função personalizada sem interface do usuário que, por exemplo, armazena um `storeValue` valor do usuário. Você pode modificá-la para que seja armazenado qualquer valor de token que você precise.

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

Os Suplementos do Office são baseados na Web e você pode usar qualquer técnica de autenticação da Web. Não há nenhum padrão ou método específico que você deve seguir para implementar sua própria autenticação com funções personalizadas sem interface do usuário. Você pode querer consultar a documentação sobre vários padrões de autenticação, começando com [este artigo sobre a autorização por serviços externos](../develop/auth-external-add-ins.md).  

Evite usar os seguintes locais para armazenar dados ao desenvolver funções personalizadas:  

- `localStorage`: funções personalizadas sem interface do usuário não têm acesso ao objeto global e, portanto, não têm acesso `window` aos dados armazenados `localStorage` em .
- `Office.context.document.settings`: Esse local não é seguro, e informações podem ser extraídas por qualquer pessoa usando o suplemento.

## <a name="dialog-box-api-example"></a>Exemplo de API da caixa de diálogo

No exemplo de código a seguir, a função usa a `getTokenViaDialog` função da API para exibir uma caixa de `Dialog` `displayWebDialogOptions` diálogo. Este exemplo é fornecido para mostrar os recursos do `Dialog` objeto, não demonstrar como autenticar.

```JavaScript
/**
 * Function retrieves a cached token or opens a dialog box if there is no saved token. Note that this is not a sufficient example of authentication but is intended to show the capabilities of the Dialog object.
 * @param {string} url URL for a stored token.
 */
function getTokenViaDialog(url) {
  return new Promise (function (resolve, reject) {
    if (_dialogOpen) {
      // Can only have one dialog box open at once. Wait for previous dialog box's token.
      let timeout = 5;
      let count = 0;
      var intervalId = setInterval(function () {
        count++;
        if(_cachedToken) {
          resolve(_cachedToken);
          clearInterval(intervalId);
        }
        if(count >= timeout) {
          reject("Timeout while waiting for token");
          clearInterval(intervalId);
        }
      }, 1000);
    } else {
      _dialogOpen = true;
      OfficeRuntime.displayWebDialog(url, {
        height: '50%',
        width: '50%',
        onMessage: function (message, dialog) {
          _cachedToken = message;
          resolve(message);
          dialog.close();
          return;
        },
        onRuntimeError: function(error, dialog) {
          reject(error);
        },
      }).catch(function (e) {
        reject(e);
      });
    }
  });
}
```

## <a name="next-steps"></a>Próximas etapas
Saiba como [depurar funções personalizadas sem](custom-functions-debugging.md)interface do usuário.

## <a name="see-also"></a>Confira também

* [Tempo de execução para funções personalizadas do Excel sem interface do usuário](custom-functions-runtime.md)
* [Tutorial de funções personalizadas do Excel](../tutorials/excel-tutorial-create-custom-functions.md)