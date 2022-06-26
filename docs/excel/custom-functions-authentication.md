---
ms.date: 06/15/2022
description: Autenticar usuários usando funções personalizadas que não usam um runtime compartilhado.
title: Autenticação para funções personalizadas sem um runtime compartilhado
ms.localizationpriority: medium
ms.openlocfilehash: 0f4493f9cf68236a9d9d83ebd3299c9ce3371560
ms.sourcegitcommit: d8fbe472b35c758753e5d2e4b905a5973e4f7b52
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 06/25/2022
ms.locfileid: "66229677"
---
# <a name="authentication-for-custom-functions-without-a-shared-runtime"></a>Autenticação para funções personalizadas sem um runtime compartilhado

Em alguns cenários, uma função personalizada que não usa um runtime compartilhado precisará autenticar o usuário para acessar recursos protegidos. Funções personalizadas que não usam uma execução de runtime compartilhado em um runtime somente javaScript. Por isso, se o suplemento tiver um painel de tarefas, você precisará passar dados entre o runtime somente JavaScript e o runtime de suporte a HTML usado pelo painel de tarefas. Faça isso usando o objeto [OfficeRuntime.storage](/javascript/api/office-runtime/officeruntime.storage) e uma API de Caixa de Diálogo especial.

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

[!include[Shared runtime note](../includes/shared-runtime-note.md)]

## <a name="officeruntimestorage-object"></a>Objeto OfficeRuntime.storage

O runtime somente JavaScript não tem `localStorage` um objeto disponível na janela global, em que você normalmente armazena dados. Em vez disso, seu código deve compartilhar dados entre funções personalizadas e painéis de tarefas usando para `OfficeRuntime.storage` definir e obter dados.

### <a name="suggested-usage"></a>Uso sugerido

Quando você precisa se autenticar de um suplemento de função personalizada que não usa um runtime compartilhado, `OfficeRuntime.storage` seu código deve verificar se o token de acesso já foi adquirido. Caso contrário, use [OfficeRuntime.displayWebDialog](/javascript/api/office-runtime#office-runtime-officeruntime-displaywebdialog-function(1)) para autenticar o usuário, recuperar o token de acesso e, em seguida, armazenar o token `OfficeRuntime.storage` para uso futuro.

## <a name="dialog-api"></a>API de caixa de diálogo

Se um token não existir, você deverá usar a `OfficeRuntime.dialog` API para solicitar que o usuário entre. Depois que um usuário insere suas credenciais, o token de acesso resultante pode ser armazenado como um item em `OfficeRuntime.storage`.

> [!NOTE]
> O runtime somente JavaScript usa um objeto de caixa de diálogo ligeiramente diferente do objeto de diálogo no runtime do mecanismo do navegador usado pelos painéis de tarefas. Ambos são chamados de "API de Caixa de Diálogo", mas usam [OfficeRuntime.displayWebDialog](/javascript/api/office-runtime#office-runtime-officeruntime-displaywebdialog-function(1)) para autenticar usuários no runtime somente JavaScript *, não* [Office.ui.displayDialogAsync](/javascript/api/office/office.ui#office-office-ui-displaydialogasync-member(1)).

O diagrama a seguir descreve esse processo básico. A linha pontilhada indica que as funções personalizadas e o painel de tarefas do suplemento fazem parte do suplemento como um todo, embora usem runtimes separados.

1. As chamadas de função personalizada de uma célula são emitdas por você em uma pasta de trabalho do Excel.
2. A função personalizada usa `OfficeRuntime.dialog` para passar suas credenciais de usuário para um site.
3. Esse site, em seguida, retorna um token de acesso para a função personalizada.
4. Em seguida, sua função personalizada define esse token de acesso para um item no `OfficeRuntime.storage`.
5. O painel de tarefas do seu suplemento acessa o token a partir de `OfficeRuntime.storage`.

![Diagrama de função personalizada usando a API da caixa de diálogo para obter o token de acesso e, em seguida, compartilhe o token com o painel de tarefas por meio da API OfficeRuntime.storage.](../images/authentication-diagram.png "Diagrama de autenticação.")

## <a name="storing-the-token"></a>Armazenando o token

Os exemplos a seguir são do exemplo de código [Usando OfficeRuntime.storage em funções personalizadas](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Excel-custom-functions/AsyncStorage). Consulte este exemplo de código para obter um exemplo completo de compartilhamento de dados entre funções personalizadas e o painel de tarefas em suplementos que não usam um runtime compartilhado.

Se a função personalizada for autenticada, ela receberá o token de acesso e precisará armazená-lo em `OfficeRuntime.storage`. O exemplo de código a seguir mostra como chamar o método `storage.setItem` para armazenar um valor. A `storeValue` função é uma função personalizada que armazena um valor do usuário. Você pode modificá-la para que seja armazenado qualquer valor de token que você precise.

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

Quando o painel de tarefas precisa do token de acesso, ele pode recuperar o token do `OfficeRuntime.storage` item. O exemplo de código a seguir mostra como usar o método `storage.getItem` para recuperar o token.

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

Os Suplementos do Office são baseados na Web e você pode usar qualquer técnica de autenticação da Web. Não há um padrão ou método específico que você deva seguir para implementar sua própria autenticação com funções personalizadas. Você pode querer consultar a documentação sobre vários padrões de autenticação, começando com [este artigo sobre a autorização por serviços externos](../develop/auth-external-add-ins.md).  

Evite usar os seguintes locais para armazenar dados ao desenvolver funções personalizadas:

- `localStorage`: funções personalizadas que não usam um runtime compartilhado não têm acesso ao objeto global `window` e, portanto, não têm acesso aos dados armazenados em `localStorage`.
- `Office.context.document.settings`: esse local não é seguro e as informações podem ser extraídas por qualquer pessoa que use o suplemento.

## <a name="dialog-box-api-example"></a>Exemplo de API da caixa de diálogo

No exemplo de código a seguir, a função `getTokenViaDialog` usa a `OfficeRuntime.displayWebDialog` função para exibir uma caixa de diálogo. Este exemplo é fornecido para mostrar os recursos do método, não demonstrar como autenticar.

```JavaScript
/**
 * Function retrieves a cached token or opens a dialog box if there is no saved token. Note that this isn't a sufficient example of authentication but is intended to show the capabilities of the displayWebDialog method.
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

Saiba como [depurar funções personalizadas](custom-functions-debugging.md).

## <a name="see-also"></a>Confira também

* [Runtime somente javaScript para funções personalizadas](custom-functions-runtime.md)
* [Tutorial de funções personalizadas do Excel](../tutorials/excel-tutorial-create-custom-functions.md)