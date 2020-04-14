---
ms.date: 04/13/2020
description: Entenda os principais cenários de desenvolvimento de funções personalizadas do Excel que usam o novo tempo de execução do JavaScript.
title: Tempo de execução de funções personalizadas do Excel
localization_priority: Normal
ms.openlocfilehash: dc049aa681ae4f7664d5bd92f925e7566c0d7103
ms.sourcegitcommit: 118e8bcbcfb73c93e2053bda67fe8dd20799b170
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 04/13/2020
ms.locfileid: "43241039"
---
# <a name="runtime-for-excel-custom-functions"></a>Tempo de execução de funções personalizadas do Excel

Funções personalizadas usam um novo tempo de execução do JavaScript, diferente do tempo de execução usado por outras partes de um suplemento, como o painel de tarefas ou outros elementos da interface do usuário. Esse tempo de execução do JavaScript foi projetado para otimizar o desempenho de cálculos em funções personalizadas, e expõe as novas APIs disponíveis para executar ações comuns baseadas na Web, dentro de funções personalizadas, como solicitação de dados externos ou troca de dados por meio de uma conexão persistente com um servidor.

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

O tempo de execução do JavaScript também fornece acesso às novas APIs no namespace `OfficeRuntime` que pode ser usado em funções personalizadas ou por outras partes de um suplemento para armazenar dados ou exibir uma caixa de diálogo. Este artigo mostra como usar essas APIs em funções personalizadas e descreve considerações adicionais para o desenvolvimento de funções personalizadas.

## <a name="requesting-external-data"></a>Como solicitar dados externos

É possível solicitar dados externos em uma função personalizada por meio de uma API, como a API [Fetch](https://developer.mozilla.org/en-US/docs/Web/API/Fetch_API), ou por meio de um objeto [XmlHttpRequest (XHR)](https://developer.mozilla.org/en-US/docs/Web/API/XMLHttpRequest), uma API Web padrão que envia solicitações HTTP para interagir com os servidores.

Dentro do tempo de execução do JavaScript usado por funções personalizadas, o XHR implementa medidas de segurança adicionais exigindo a [mesma política de origem](https://developer.mozilla.org/en-US/docs/Web/Security/Same-origin_policy) e o [CORS](https://www.w3.org/TR/cors/)simples.

Observe que uma implementação CORS simples não pode usar cookies e é compatível apenas com métodos simples (GET, HEAD, POST). A CORS simples aceita cabeçalhos simples com nomes de campos `Accept`, `Accept-Language`, `Content-Language`. Você também pode usar um `Content-Type` cabeçalho no CORS simples, desde que o tipo de conteúdo `application/x-www-form-urlencoded`seja `text/plain`, ou `multipart/form-data`.

### <a name="xhr-example"></a>Exemplo de XHR

No código de exemplo a seguir, a função `getTemperature` chama a função `sendWebRequest` para obter a temperatura de uma área específica, de acordo com a ID do termômetro. A função `sendWebRequest` usa XHR para emitir uma solicitação `GET` para um ponto de extremidade que fornece os dados.

> [!NOTE] 
> Se usar fetch ou XHR, uma nova `Promise` JavaScript será retornada. Antes de setembro de 2018, era necessário especificar `OfficeExtension.Promise` para usar promessas na API JavaScript para Office, mas agora, basta usar um `Promise` JavaScript.

```js
function getTemperature(thermometerID) {
  return new Promise(function(setResult) {
      sendWebRequest(thermometerID, function(data){ 
          storeLastTemperature(thermometerID, data.temperature);
          setResult(data.temperature);
      });
  });
}

// Helper method that uses Office's implementation of XMLHttpRequest in the JavaScript runtime for custom functions  
function sendWebRequest(thermometerID, data) {
    var xhttp = new XMLHttpRequest();
    xhttp.onreadystatechange = function() {
        if (this.readyState == 4 && this.status == 200) {
           data.temperature = JSON.parse(xhttp.responseText).temperature
        };
        
        //set Content-Type to application/text. Application/json is not currently supported with Simple CORS
        xhttp.setRequestHeader("Content-Type", "application/text");
        xhttp.open("GET", "https://contoso.com/temperature/" + thermometerID), true)
        xhttp.send();  
    }
}
```

## <a name="receiving-data-via-websockets"></a>Como receber dados por meio de WebSockets

Em uma função personalizada, é possível usar [WebSockets](https://developer.mozilla.org/en-US/docs/Web/API/WebSockets_API) para trocar dados por meio de uma conexão persistente com um servidor. Usando WebSockets, a função personalizada pode abrir uma conexão com um servidor e, em seguida, receber mensagens do servidor automaticamente, quando determinados eventos ocorrerem, sem precisar consultar explicitamente os dados do servidor.

### <a name="websockets-example"></a>Exemplo de WebSockets

O código de exemplo a seguir estabelece uma conexão `WebSocket` e registra cada mensagem de entrada do servidor.

```js
const ws = new WebSocket('wss://bundles.office.com');
ws.onmessage = function (message) {
    console.log(`Received: ${message}`);
}
ws.onerror = function (error) {
    console.err(`Failed: ${error}`);
}
```

## <a name="storing-and-accessing-data"></a>Como armazenar e acessar os dados

Em uma função personalizada (ou em outras partes de um suplemento), você pode armazenar e acessar dados usando o objeto `OfficeRuntime.storage`. `Storage` é um sistema de armazenamento de chave-valor persistente e não criptografado, que fornece uma alternativa para [localStorage](https://developer.mozilla.org/en-US/docs/Web/API/Window/localStorage), que não pode ser usado em funções personalizadas. `Storage`o oferece 10 MB de dados por domínio. Os domínios podem ser compartilhados por mais de um suplemento.

`Storage` é uma solução de armazenamento compartilhado, o que significa que várias partes de um suplemento podem acessar os mesmos dados. Por exemplo, tokens para autenticação de usuário podem ser armazenados em `storage`, já que ele pode ser acessado tanto por uma função personalizada quanto por elementos da interface do usuário de um suplemento, como um painel de tarefas. Da mesma forma, se dois suplementos compartilham o mesmo domínio (por exemplo, `www.contoso.com/addin1` `www.contoso.com/addin2`), eles também podem compartilhar informações de frente e para trás `storage`. Observe que os suplementos que possuem subdomínios diferentes terão instâncias diferentes `storage` (por exemplo, `subdomain.contoso.com/addin1`, `differentsubdomain.contoso.com/addin2`).

Como `storage` pode ser um local compartilhado, é importante notar que é possível substituir os pares chave-valor.

Os métodos a seguir estão disponíveis no objeto `storage`:

 - `getItem`
 - `getItems`
 - `setItem`
 - `setItems`
 - `removeItem`
 - `removeItems`
 - `getKeys`

.[!NOTE]
> Não há nenhum método para limpar todas as informações (como `clear`). Em vez disso, use `removeItems` para remover várias entradas de uma só vez.

### <a name="officeruntimestorage-example"></a>Exemplo de OfficeRuntime. Storage

O exemplo de código a seguir `OfficeRuntime.storage.setItem` chama a função para definir uma chave e `storage`um valor para.

```js
function StoreValue(key, value) {

  return OfficeRuntime.storage.setItem(key, value).then(function (result) {
      return "Success: Item with key '" + key + "' saved to storage.";
  }, function (error) {
      return "Error: Unable to save item with key '" + key + "' to storage. " + error;
  });
}
```

## <a name="additional-considerations"></a>Considerações adicionais

Para criar um suplemento que será executado em várias plataformas (um dos principais locatários de Suplementos do Office), você não deve acessar o DOM (Modelo de Objeto do Documento) em funções personalizadas nem usar bibliotecas, como a jQuery, que dependem do DOM. No Excel no Windows, onde as funções personalizadas usam o tempo de execução do JavaScript, as funções personalizadas não podem acessar o DOM.

## <a name="next-steps"></a>Próximas etapas
Saiba como [realizar solicitações da Web com funções personalizadas](custom-functions-web-reqs.md).

## <a name="see-also"></a>Confira também

* [Criar funções personalizadas no Excel](custom-functions-overview.md)
* [Arquitetura de funções personalizadas](custom-functions-architecture.md)
* [Exibir uma caixa de diálogo em funções personalizadas](custom-functions-dialog.md)
* [Tutorial de funções personalizadas](../tutorials/excel-tutorial-create-custom-functions.md)
