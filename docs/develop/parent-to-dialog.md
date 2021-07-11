---
title: Maneiras alternativas de passar mensagens para uma caixa de diálogo de sua página host
description: Saiba soluções alternativas a ser usadas quando o método messageChild não é suportado.
ms.date: 09/24/2020
localization_priority: Normal
ms.openlocfilehash: 8da6bc3e1231bc6296a16fa153dc0e4ba1bd102b
ms.sourcegitcommit: 883f71d395b19ccfc6874a0d5942a7016eb49e2c
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 07/09/2021
ms.locfileid: "53349774"
---
# <a name="alternative-ways-of-passing-messages-to-a-dialog-box-from-its-host-page"></a><span data-ttu-id="943fa-103">Maneiras alternativas de passar mensagens para uma caixa de diálogo de sua página host</span><span class="sxs-lookup"><span data-stu-id="943fa-103">Alternative ways of passing messages to a dialog box from its host page</span></span>

<span data-ttu-id="943fa-104">A maneira recomendada de passar dados e mensagens de uma página pai para uma caixa de diálogo filho é com o método conforme descrito em Usar a API de diálogo Office em seus `messageChild` [Office Desempis.](dialog-api-in-office-add-ins.md#pass-information-to-the-dialog-box) Se o seu add-in estiver em execução em uma plataforma ou host que não oferece suporte ao conjunto de requisitos [DialogApi 1.2,](../reference/requirement-sets/dialog-api-requirement-sets.md)há duas outras maneiras de passar informações para a caixa de diálogo:</span><span class="sxs-lookup"><span data-stu-id="943fa-104">The recommended way to pass data and messages from a parent page to a child dialog box is with the `messageChild` method as described in [Use the Office dialog API in your Office Add-ins](dialog-api-in-office-add-ins.md#pass-information-to-the-dialog-box). If your add-in is running on a platform or host that does not support the [DialogApi 1.2 requirement set](../reference/requirement-sets/dialog-api-requirement-sets.md), there are two other ways that you can pass information to the dialog box:</span></span>

- <span data-ttu-id="943fa-105">Adicionar parâmetros de consulta à URL que é transmitida para `displayDialogAsync`.</span><span class="sxs-lookup"><span data-stu-id="943fa-105">Add query parameters to the URL that is passed to `displayDialogAsync`.</span></span>
- <span data-ttu-id="943fa-106">Armazenar as informações em outro local que seja acessível para a janela do host e para a caixa de diálogo.</span><span class="sxs-lookup"><span data-stu-id="943fa-106">Store the information somewhere that is accessible to both the host window and dialog box.</span></span> <span data-ttu-id="943fa-107">As duas janelas não compartilham um armazenamento de sessão comum (a propriedade [Window.sessionStorage),](https://developer.mozilla.org/docs/Web/API/Window/sessionStorage) mas se elas têm o mesmo domínio *(incluindo* o número da porta, se for o caso), elas compartilham um local [comum Armazenamento](https://www.w3schools.com/html/html5_webstorage.asp).\*</span><span class="sxs-lookup"><span data-stu-id="943fa-107">The two windows do not share a common session storage (the [Window.sessionStorage](https://developer.mozilla.org/docs/Web/API/Window/sessionStorage) property), but *if they have the same domain* (including port number, if any), they share a common [Local Storage](https://www.w3schools.com/html/html5_webstorage.asp).\*</span></span>


> [!NOTE]
> <span data-ttu-id="943fa-108">\* Há um bug que afetará sua estratégia de tratamento de tokens.</span><span class="sxs-lookup"><span data-stu-id="943fa-108">\* There is a bug that will effect your strategy for token handling.</span></span> <span data-ttu-id="943fa-109">Se o suplemento estiver sendo executado no **Office na Web** nos navegadores Safari ou Microsoft Edge, o painel de tarefas e a caixa de diálogo não compartilharão o mesmo Armazenamento Local, portanto, ele não poderá ser usado para a comunicação entre eles.</span><span class="sxs-lookup"><span data-stu-id="943fa-109">If the add-in is running in **Office on the web** in either the Safari or Edge browser, the dialog box and task pane do not share the same Local Storage, so it cannot be used to communicate between them.</span></span>

## <a name="use-local-storage"></a><span data-ttu-id="943fa-110">Usar o armazenamento local</span><span class="sxs-lookup"><span data-stu-id="943fa-110">Use local storage</span></span>

<span data-ttu-id="943fa-111">Para usar o armazenamento local, chame o método do objeto na página host antes da `setItem` `window.localStorage` `displayDialogAsync` chamada, como no exemplo a seguir.</span><span class="sxs-lookup"><span data-stu-id="943fa-111">To use local storage, call the `setItem` method of the `window.localStorage` object in the host page before the `displayDialogAsync` call, as in the following example.</span></span>

```js
localStorage.setItem("clientID", "15963ac5-314f-4d9b-b5a1-ccb2f1aea248");
```

<span data-ttu-id="943fa-112">O código na caixa de diálogo lê o item quando necessário, como no exemplo a seguir.</span><span class="sxs-lookup"><span data-stu-id="943fa-112">Code in the dialog box reads the item when it's needed, as in the following example.</span></span>

```js
var clientID = localStorage.getItem("clientID");
// You can also use property syntax:
// var clientID = localStorage.clientID;
```

## <a name="use-query-parameters"></a><span data-ttu-id="943fa-113">Usar os parâmetros de consulta</span><span class="sxs-lookup"><span data-stu-id="943fa-113">Use query parameters</span></span>

<span data-ttu-id="943fa-114">O exemplo a seguir mostra como transmitir dados com um parâmetro de consulta.</span><span class="sxs-lookup"><span data-stu-id="943fa-114">The following example shows how to pass data with a query parameter.</span></span>

```js
Office.context.ui.displayDialogAsync('https://myAddinDomain/myDialog.html?clientID=15963ac5-314f-4d9b-b5a1-ccb2f1aea248');
```

<span data-ttu-id="943fa-115">Para ver um exemplo que usa essa técnica, consulte [Inserir gráficos do Excel usando o Microsoft Graph em um Suplemento do PowerPoint](https://github.com/OfficeDev/PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart).</span><span class="sxs-lookup"><span data-stu-id="943fa-115">For a sample that uses this technique, see [Insert Excel charts using Microsoft Graph in a PowerPoint add-in](https://github.com/OfficeDev/PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart).</span></span>

<span data-ttu-id="943fa-116">O código na caixa de diálogo pode analisar a URL e ler o valor do parâmetro.</span><span class="sxs-lookup"><span data-stu-id="943fa-116">Code in your dialog box can parse the URL and read the parameter value.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="943fa-117">Office adiciona automaticamente um parâmetro de consulta `_host_info` chamado à URL que é passada para `displayDialogAsync` .</span><span class="sxs-lookup"><span data-stu-id="943fa-117">Office automatically adds a query parameter called `_host_info` to the URL that is passed to `displayDialogAsync`.</span></span> <span data-ttu-id="943fa-118">(Ele é anexado após os parâmetros de consulta personalizados, se algum.</span><span class="sxs-lookup"><span data-stu-id="943fa-118">(It is appended after your custom query parameters, if any.</span></span> <span data-ttu-id="943fa-119">Não é anexado a URLs subsequentes para as que a caixa de diálogo navega.) A Microsoft pode alterar o conteúdo desse valor ou removê-lo completamente, no futuro, para que seu código não o leia.</span><span class="sxs-lookup"><span data-stu-id="943fa-119">It is not appended to any subsequent URLs that the dialog box navigates to.) Microsoft may change the content of this value, or remove it entirely, in the future, so your code should not read it.</span></span> <span data-ttu-id="943fa-120">O mesmo valor é adicionado ao armazenamento de sessão da caixa de diálogo (a [propriedade Window.sessionStorage).](https://developer.mozilla.org/docs/Web/API/Window/sessionStorage)</span><span class="sxs-lookup"><span data-stu-id="943fa-120">The same value is added to the dialog box's session storage (the [Window.sessionStorage](https://developer.mozilla.org/docs/Web/API/Window/sessionStorage) property).</span></span> <span data-ttu-id="943fa-121">Novamente, *seu código não deve ler nem gravar nesse valor*.</span><span class="sxs-lookup"><span data-stu-id="943fa-121">Again, *your code should neither read nor write to this value*.</span></span>
