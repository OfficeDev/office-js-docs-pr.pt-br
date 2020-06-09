---
ms.date: 05/17/2020
description: Autenticar usuários usando funções personalizadas no Excel que não usam o painel de tarefas.
title: Autenticação para funções personalizadas sem interface do usuário
localization_priority: Normal
ms.openlocfilehash: b4ff234f71ed2a36cc311e45f47498d19380b862
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 06/08/2020
ms.locfileid: "44609335"
---
# <a name="authentication-for-ui-less-custom-functions"></a>Autenticação para funções personalizadas sem interface do usuário

Em alguns cenários, a função personalizada que não usa um painel de tarefas ou outros elementos de interface do usuário (função menos personalizada) precisará autenticar o usuário para acessar recursos protegidos. Esteja ciente de que as funções personalizadas sem interface do usuário são executadas em um tempo de execução do JavaScript. Por causa disso, você precisará transmitir dados entre o tempo de execução do JavaScript somente e o tempo de execução típico do mecanismo de navegador usado pela maioria dos suplementos usando o `OfficeRuntime.storage` objeto e a API da caixa de diálogo.

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

[!include[Shared runtime note](../includes/shared-runtime-note.md)]

## <a name="officeruntimestorage-object"></a>Objeto OfficeRuntime.storage

O tempo de execução do JavaScript somente usado por funções personalizadas sem interface do usuário não tem um `localStorage` objeto disponível na janela global, onde você normalmente armazena dados. Em vez disso, você deve compartilhar dados entre funções personalizadas e painéis de tarefas sem interface do usuário usando o [OfficeRuntime. Storage](/javascript/api/office-runtime/officeruntime.storage) para definir e obter dados.

### <a name="suggested-usage"></a>Uso sugerido

Quando você precisar autenticar a partir de uma função personalizada sem interface do usuário, verifique `storage` se o token de acesso já foi adquirido. Caso contrário, use a API de caixa de diálogo para autenticar o usuário, recuperar o token de acesso e, em seguida, armazenar o token em `storage` para uso futuro.

## <a name="dialog-api"></a>API de Caixa de Diálogo

Se um token não existir, você deverá usar a API de diálogo para solicitar que o usuário faça logon. Depois que um usuário insere suas credenciais, o token de acesso resultante pode ser armazenado em `storage`.

> [!NOTE]
> O tempo de execução do JavaScript somente usa um objeto Dialog que é ligeiramente diferente do objeto Dialog no tempo de execução do mecanismo do navegador usado por painéis de tarefas. Eles são conhecidos como "API da caixa de diálogo", mas usam `OfficeRuntime.Dialog` para autenticar usuários no tempo de execução do JavaScript.

O diagrama a seguir descreve esse processo básico. A linha pontilhada indica que as funções personalizadas sem interface do usuário e o painel de tarefas do suplemento fazem parte do seu suplemento como um todo, embora usem tempos de execução separados.

1. Você emite uma chamada de função personalizada sem interface do usuário a partir de uma célula em uma pasta de trabalho do Excel.
2. A função personalizada sem interface do usuário usa o `Dialog` para passar suas credenciais de usuário para um site.
3. Em seguida, este site retorna um token de acesso à função personalizada sem interface do usuário.
4. Sua função personalizada sem interface do usuário define esse token de acesso para o `storage` .
5. O painel de tarefas do seu suplemento acessa o token a partir de `storage`.

![Diagrama da função personalizada usando a API da caixa de diálogo para obter o token de acesso e compartilhar o token com o painel de tarefas por meio da API OfficeRuntime. Storage.](../images/authentication-diagram.png "Diagrama de autenticação.")

## <a name="storing-the-token"></a>Armazenando o token

Os exemplos a seguir são do exemplo de código [Usando OfficeRuntime.storage em funções personalizadas](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Excel-custom-functions/AsyncStorage). Consulte este exemplo de código para obter um exemplo completo de compartilhamento de dados entre as funções personalizadas sem interface do usuário e o painel de tarefas.

Se a função personalizada sem interface do usuário for autenticada, ela receberá o token de acesso e deverá armazená-lo no `storage` . O exemplo de código a seguir mostra como chamar o método `storage.setItem` para armazenar um valor. A `storeValue` função é uma função personalizada sem IU que, por exemplo, armazena um valor do usuário. Você pode modificá-la para que seja armazenado qualquer valor de token que você precise.

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

Os Suplementos do Office são baseados na Web e você pode usar qualquer técnica de autenticação da Web. Não há um padrão ou método específico que você deve seguir para implementar sua própria autenticação com funções personalizadas sem interface do usuário. Você pode querer consultar a documentação sobre vários padrões de autenticação, começando com [este artigo sobre a autorização por serviços externos](../develop/auth-external-add-ins.md).  

Evite usar os seguintes locais para armazenar dados ao desenvolver funções personalizadas:  

- `localStorage`: As funções personalizadas sem interface do usuário não têm acesso ao `window` objeto global e, portanto, não têm acesso aos dados armazenados no `localStorage` .
- `Office.context.document.settings`: Esse local não é seguro, e informações podem ser extraídas por qualquer pessoa usando o suplemento.

## <a name="dialog-box-api-example"></a>Exemplo de API da caixa de diálogo

No exemplo de código a seguir, a função `getTokenViaDialog` usa a `Dialog` função da API `displayWebDialogOptions` para exibir uma caixa de diálogo. Este exemplo é fornecido para mostrar os recursos do `Dialog` objeto, não demonstra como autenticar.

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
Saiba como [depurar funções personalizadas sem interface do usuário](custom-functions-debugging.md).

## <a name="see-also"></a>Confira também

* [Tempo de execução para funções personalizadas do Excel sem interface do usuário](custom-functions-runtime.md)
* [Tutorial de funções personalizadas do Excel](excel-tutorial-custom-functions.md)
