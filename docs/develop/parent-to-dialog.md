---
title: Maneiras alternativas de passar mensagens para uma caixa de diálogo de sua página host
description: Saiba soluções alternativas a ser usadas quando o método messageChild não é suportado.
ms.date: 07/08/2021
ms.localizationpriority: medium
ms.openlocfilehash: 94b002970084a13fe69e1469aa26577039600989
ms.sourcegitcommit: 287a58de82a09deeef794c2aa4f32280efbbe54a
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/28/2022
ms.locfileid: "64496807"
---
# <a name="alternative-ways-of-passing-messages-to-a-dialog-box-from-its-host-page"></a>Maneiras alternativas de passar mensagens para uma caixa de diálogo de sua página host

A maneira recomendada `messageChild` de passar dados e mensagens de uma página pai para uma caixa de diálogo filho é com o método conforme descrito em Use a API de diálogo Office em seus [Office Add-ins](dialog-api-in-office-add-ins.md#pass-information-to-the-dialog-box). Se o seu add-in estiver em execução em uma plataforma ou host que não oferece suporte ao conjunto de requisitos [dialogApi 1.2](/javascript/api/requirement-sets/common/dialog-api-requirement-sets), há duas outras maneiras de passar informações para a caixa de diálogo.

- Adicionar parâmetros de consulta à URL que é transmitida para `displayDialogAsync`.
- Armazenar as informações em outro local que seja acessível para a janela do host e para a caixa de diálogo. As duas janelas não compartilham um armazenamento de sessão comum (a propriedade [Window.sessionStorage](https://developer.mozilla.org/docs/Web/API/Window/sessionStorage)), mas se elas têm o mesmo *domínio (incluindo* o número da porta, se for o caso), elas compartilham um local [comum](https://www.w3schools.com/html/html5_webstorage.asp) Armazenamento.\*

> [!NOTE]
> \* Há um bug que afetará sua estratégia de tratamento de tokens. Se o suplemento estiver sendo executado no **Office na Web** nos navegadores Safari ou Microsoft Edge, o painel de tarefas e a caixa de diálogo não compartilharão o mesmo Armazenamento Local, portanto, ele não poderá ser usado para a comunicação entre eles.

## <a name="use-local-storage"></a>Usar o armazenamento local

Para usar o armazenamento local, chame `setItem` o método do `window.localStorage` objeto na página host antes `displayDialogAsync` da chamada, como no exemplo a seguir.

```js
localStorage.setItem("clientID", "15963ac5-314f-4d9b-b5a1-ccb2f1aea248");
```

O código na caixa de diálogo lê o item quando necessário, como no exemplo a seguir.

```js
var clientID = localStorage.getItem("clientID");
// You can also use property syntax:
// var clientID = localStorage.clientID;
```

## <a name="use-query-parameters"></a>Usar os parâmetros de consulta

O exemplo a seguir mostra como transmitir dados com um parâmetro de consulta.

```js
Office.context.ui.displayDialogAsync('https://myAddinDomain/myDialog.html?clientID=15963ac5-314f-4d9b-b5a1-ccb2f1aea248');
```

Para ver um exemplo que usa essa técnica, consulte [Inserir gráficos do Excel usando o Microsoft Graph em um Suplemento do PowerPoint](https://github.com/OfficeDev/PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart).

O código na caixa de diálogo pode analisar a URL e ler o valor do parâmetro.

> [!IMPORTANT]
> Office adiciona automaticamente um parâmetro de consulta chamado `_host_info` à URL que é passada para `displayDialogAsync`. (Ele é anexado após os parâmetros de consulta personalizados, se algum. Não é anexado a URLs subsequentes para as que a caixa de diálogo navega.) A Microsoft pode alterar o conteúdo desse valor ou removê-lo completamente, no futuro, para que seu código não o leia. O mesmo valor é adicionado ao armazenamento de sessão da caixa de diálogo (a [propriedade Window.sessionStorage](https://developer.mozilla.org/docs/Web/API/Window/sessionStorage) ). Novamente, *seu código não deve ler nem gravar nesse valor*.
