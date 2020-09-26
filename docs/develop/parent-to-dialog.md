---
title: Maneiras alternativas de passar mensagens para uma caixa de diálogo da página host
description: Saiba como solucionar contornos para usar quando não há suporte para o método messageChild.
ms.date: 09/24/2020
localization_priority: Normal
ms.openlocfilehash: 8f44f7f5c145b58d13e7387d01e28fd349a512fc
ms.sourcegitcommit: b47318a24a50443b0579e05e178b3bb5433c372f
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 09/25/2020
ms.locfileid: "48279480"
---
# <a name="alternative-ways-of-passing-messages-to-a-dialog-box-from-its-host-page"></a>Maneiras alternativas de passar mensagens para uma caixa de diálogo da página host

A maneira recomendada de passar dados e mensagens de uma página pai para uma caixa de diálogo filha é com o `messageChild` método conforme descrito em [usar a API de caixa de diálogo do Office em seus suplementos do Office](dialog-api-in-office-add-ins.md#pass-information-to-the-dialog-box). Se o suplemento estiver sendo executado em uma plataforma ou host que não ofereça suporte ao [conjunto de requisitos DialogApi 1,2](../reference/requirement-sets/dialog-api-requirement-sets.md), há duas outras maneiras de passar informações para a caixa de diálogo:

- Adicionar parâmetros de consulta à URL que é transmitida para `displayDialogAsync`.
- Armazenar as informações em outro local que seja acessível para a janela do host e para a caixa de diálogo. As duas janelas não compartilham um armazenamento de sessão comum (a propriedade [Window. sessionStorage](https://developer.mozilla.org/docs/Web/API/Window/sessionStorage) ), mas *se tiverem o mesmo domínio* (incluindo o número da porta, se houver), eles compartilharão um [armazenamento local](https://www.w3schools.com/html/html5_webstorage.asp)comum.\*


> [!NOTE]
> \* Há um bug que afetará sua estratégia de tratamento de tokens. Se o suplemento estiver sendo executado no **Office na Web** nos navegadores Safari ou Microsoft Edge, o painel de tarefas e a caixa de diálogo não compartilharão o mesmo Armazenamento Local, portanto, ele não poderá ser usado para a comunicação entre eles.

## <a name="use-local-storage"></a>Usar o armazenamento local

Para usar o armazenamento local, chame o `setItem` método do `window.localStorage` objeto na página host antes da `displayDialogAsync` chamada, como no exemplo a seguir:

```js
localStorage.setItem("clientID", "15963ac5-314f-4d9b-b5a1-ccb2f1aea248");
```

O código na caixa de diálogo lê o item quando necessário, como no exemplo a seguir:

```js
var clientID = localStorage.getItem("clientID");
// You can also use property syntax:
// var clientID = localStorage.clientID;
```

## <a name="use-query-parameters"></a>Usar parâmetros de consulta

O exemplo a seguir mostra como transmitir dados com um parâmetro de consulta:

```js
Office.context.ui.displayDialogAsync('https://myAddinDomain/myDialog.html?clientID=15963ac5-314f-4d9b-b5a1-ccb2f1aea248');
```

Para ver um exemplo que usa essa técnica, consulte [Inserir gráficos do Excel usando o Microsoft Graph em um Suplemento do PowerPoint](https://github.com/OfficeDev/PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart).

O código na caixa de diálogo pode analisar a URL e ler o valor do parâmetro.

> [!IMPORTANT]
> O Office adiciona automaticamente um parâmetro de consulta chamado `_host_info` à URL que é transmitida para `displayDialogAsync` . (Ele é acrescentado após os parâmetros de consulta personalizados, se houver. Ele não é acrescentado a quaisquer URLs subsequentes para as quais a caixa de diálogo navega.) A Microsoft pode alterar o conteúdo desse valor ou removê-lo totalmente, no futuro, portanto, o código não deve lê-lo. O mesmo valor é adicionado ao armazenamento da sessão da caixa de diálogo (a propriedade [Window. sessionStorage](https://developer.mozilla.org/docs/Web/API/Window/sessionStorage) ). Novamente, *seu código não deve ser lido nem gravado para esse valor*.
