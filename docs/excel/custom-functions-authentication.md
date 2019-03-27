---
ms.date: 03/19/2019
description: Autenticar usuários usando funções personalizadas no Excel.
title: Autenticação para funções personalizadas
ms.openlocfilehash: 7db46e40758ea0282a2fd7c4d40739304a874e76
ms.sourcegitcommit: a2950492a2337de3180b713f5693fe82dbdd6a17
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/27/2019
ms.locfileid: "30871490"
---
# <a name="authentication"></a>Autenticação

Em alguns cenários, a função personalizada precisará autenticar o usuário para poder acessar recursos protegidos. Embora as funções personalizadas não exijam um método de autenticação específico, você deve estar ciente de que as funções personalizadas são executadas em um tempo de execução separado do painel de tarefas e de outros elementos de interface do usuário do seu suplemento. Por causa disso, você precisará transmitir dados entre os dois tempos de execução usando o `AsyncStorage` objeto e a API da caixa de diálogo.
  
## <a name="asyncstorage-object"></a>Objeto AsyncStorage

O tempo de execução de funções personalizadas `localStorage` não tem um objeto disponível na janela global, onde você normalmente pode armazenar dados. Em vez disso, você deve compartilhar dados entre funções personalizadas e painéis de tarefas usando o [OfficeRuntime. AsyncStorage](/javascript/api/office-runtime/officeruntime.asyncstorage) para definir e obter dados.

Além disso, há um benefício em usar `AsyncStorage`o; Ele usa um ambiente de área restrita seguro para que seus dados não possam ser acessados por outros suplementos.

### <a name="suggested-usage"></a>Uso sugerido

Quando você precisar autenticar do painel de tarefas ou de uma função personalizada, verifique `AsyncStorage` se o token de acesso já foi adquirido. Caso contrário, use a API de caixa de diálogo para autenticar o usuário, recuperar o token de acesso e armazená-lo `AsyncStorage` para uso futuro.

## <a name="dialog-api"></a>API da caixa de diálogo

Se não houver um token, você deverá usar a API da caixa de diálogo para solicitar que o usuário entre. Após um usuário inserir suas credenciais, o token de acesso resultante poderá ser armazenado `AsyncStorage`no.

> [!NOTE]
> O tempo de execução de funções personalizadas usa um objeto Dialog que é ligeiramente diferente do objeto Dialog no tempo de execução do mecanismo de navegador usado por painéis de tarefas. Eles são conhecidos como "API da caixa de diálogo", mas usam `Officeruntime.Dialog` para autenticar usuários no tempo de execução de funções personalizadas.

Para obter informações sobre como usar o `OfficeRuntime.Dialog`, consulte [Custom Functions Runtime](/office/dev/add-ins/excel/custom-functions-runtime?view=office-js#displaying-a-dialog-box).

Ao planejar todo o processo de autenticação como um todo, talvez seja útil pensar no painel de tarefas e nos elementos de interface do usuário do suplemento e das funções personalizadas, que fazem parte do suplemento como entidades separadas que podem se comunicar entre si `AsyncStorage`.

O diagrama a seguir descreve esse processo básico. Observe que a linha pontilhada indica que, enquanto elas executam ações separadas, as funções personalizadas e o painel de tarefas do seu suplemento fazem parte do seu suplemento como um todo.

1. Você emite uma chamada de função personalizada a partir de uma célula em uma pasta de trabalho do Excel.
2. A função personalizada usa `Officeruntime.Dialog` o para passar suas credenciais de usuário para um site.
3. Este site, em seguida, retorna um token de acesso para a função personalizada.
4. Sua função personalizada então define esse token de acesso para `AsyncStorage`o.
5. O painel de tarefas do suplemento acessa o token de `AsyncStorage`.

![Diagrama da função personalizada usando a API da caixa de diálogo para obter o token de acesso e compartilhar o token com o painel de tarefas por meio da API AsyncStorage.] (../images/authentication-diagram.png "Diagrama de autenticação.")

## <a name="storing-the-token"></a>Armazenar o token

Os exemplos a seguir são do [usando AsyncStorage no exemplo de código de funções personalizadas](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Excel-custom-functions/AsyncStorage) . Consulte este exemplo de código para obter um exemplo completo de compartilhamento de dados entre funções personalizadas e o painel de tarefas.

Se a função personalizada autenticar, ela receberá o token de acesso e deverá armazená-lo `AsyncStorage`no. O exemplo de código a seguir mostra como chamar `AsyncStorage.setItem` o método para armazenar um valor. A `StoreValue` função é uma função personalizada que, por exemplo, armazena um valor do usuário. Você pode modificá-lo para armazenar qualquer valor de token necessário.

```javascript
function StoreValue(key, value) {
  return OfficeRuntime.AsyncStorage.setItem(key, value).then(function (result) {
      return "Success: Item with key '" + key + "' saved to AsyncStorage.";
  }, function (error) {
      return "Error: Unable to save item with key '" + key + "' to AsyncStorage. " + error;
  });
}
```

Quando o painel de tarefas precisa do token de acesso, ele pode recuperar o `AsyncStorage`token de. O exemplo de código a seguir mostra como usar `AsyncStorage.getItem` o método para recuperar o token.

```javascript
function ReceiveTokenFromCustomFunction() {
   var key = "token";
   var tokenSendStatus = document.getElementById('tokenSendStatus');
   OfficeRuntime.AsyncStorage.getItem(key).then(function (result) {
      tokenSendStatus.value = "Success: Item with key '" + key + "' read from AsyncStorage.";
      document.getElementById('tokenTextBox2').value = result;
   }, function (error) {
      tokenSendStatus.value = "Error: Unable to read item with key '" + key + "' from AsyncStorage. " + error;
   });
}
```

## <a name="general-guidance"></a>Orientação geral

Os suplementos do Office são baseados na Web e você pode usar qualquer técnica de autenticação da Web. Não há um padrão ou método específico que você deve seguir para implementar sua própria autenticação com funções personalizadas. Você pode querer consultar a documentação sobre vários padrões de autenticação, começando com [Este artigo sobre como autorizar por meio de serviços externos](/office/dev/add-ins/develop/auth-external-add-ins?view=office-js).  

Evite usar os seguintes locais para armazenar dados ao desenvolver funções personalizadas:  

- `localStorage`: As funções personalizadas não têm acesso ao objeto global `window` e, portanto, não têm acesso aos dados armazenados `localStorage`no.
- `Office.context.document.settings`: Esse local não é seguro e as informações podem ser extraídas por qualquer pessoa que use o suplemento.

## <a name="see-also"></a>Confira também

* [Metadados de funções personalizadas](custom-functions-json.md)
* [Tempo de execução de funções personalizadas do Excel](custom-functions-runtime.md)
* [Práticas recomendadas de funções personalizadas](custom-functions-best-practices.md).
* [Tutorial de funções personalizadas do Excel](excel-tutorial-custom-functions.md)
