---
ms.date: 05/08/2019
description: Use `OfficeRuntime.storage` para salvar o estado com funções personalizadas.
title: Salvar e compartilhar o estado em funções personalizadas
localization_priority: Priority
ms.openlocfilehash: b1472b0623d15882dabff16f8be3f74756e3b3de
ms.sourcegitcommit: a99be9c4771c45f3e07e781646e0e649aa47213f
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/11/2019
ms.locfileid: "33951967"
---
## <a name="save-and-share-state-in-custom-functions"></a><span data-ttu-id="3fc0f-103">Salvar e compartilhar o estado em funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="3fc0f-103">Save and share state in custom functions</span></span>

<span data-ttu-id="3fc0f-104">Use o objeto `OfficeRuntime.storage` para salvar o estado relacionado às funções personalizadas ou o painel de tarefas no seu suplemento.</span><span class="sxs-lookup"><span data-stu-id="3fc0f-104">Use the `OfficeRuntime.storage` object to save state related to custom functions or the task pane in your add-in.</span></span> <span data-ttu-id="3fc0f-105">O armazenamento é limitado a 10 MB por domínio (que pode ser compartilhado entre vários suplementos).</span><span class="sxs-lookup"><span data-stu-id="3fc0f-105">Storage is limited to 10 MB per domain (which may be shared across multiple add-ins).</span></span> <span data-ttu-id="3fc0f-106">No Excel no Windows, o objeto `storage` é uma localização separada dentro do tempo de execução das funções personalizadas, mas no Excel Online e no Excel para Mac, o objeto `storage` é o mesmo que o `localStorage` do navegador.</span><span class="sxs-lookup"><span data-stu-id="3fc0f-106">On Excel for Windows, the `storage` object is a separate location within the custom functions runtime, but for Excel Online and Excel for Mac, the `storage` object is the same as the browser's `localStorage`.</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

<span data-ttu-id="3fc0f-107">Existem várias maneiras de usar `storage` para o gerenciamento de estado:</span><span class="sxs-lookup"><span data-stu-id="3fc0f-107">There are multiple ways to use `storage` for state management:</span></span>

- <span data-ttu-id="3fc0f-108">Você pode armazenar valores padrão para funções personalizadas para usar quando você estiver offline e não for possível acessar um recurso da Web.</span><span class="sxs-lookup"><span data-stu-id="3fc0f-108">You can store default values for custom functions to use when you are offline and unable to reach a web resource.</span></span>
- <span data-ttu-id="3fc0f-109">Você pode salvar valores para funções personalizadas para evitar fazer chamadas adicionais à um recurso da Web.</span><span class="sxs-lookup"><span data-stu-id="3fc0f-109">You can save values for custom functions to use to avoid making additional calls to a web resource.</span></span>
- <span data-ttu-id="3fc0f-110">Você pode salvar valores da sua função personalizada.</span><span class="sxs-lookup"><span data-stu-id="3fc0f-110">You can save values from your custom function.</span></span>
- <span data-ttu-id="3fc0f-111">Você pode armazenar valores do seu painel de tarefas.</span><span class="sxs-lookup"><span data-stu-id="3fc0f-111">You can store values from your task pane.</span></span>

<span data-ttu-id="3fc0f-112">O exemplo de código a seguir ilustra como armazenar um item em `storage` e recuperá-lo.</span><span class="sxs-lookup"><span data-stu-id="3fc0f-112">The following code sample illustrates how to store an item into `storage` and retrieve it.</span></span>

```js
function storeValue(key, value) {
  return OfficeRuntime.storage.setItem(key, value).then(function (result) {
      return "Success: Item with key '" + key + "' saved to storage.";
  }, function (error) {
      return "Error: Unable to save item with key '" + key + "' to storage. " + error;
  });
}

function GetValue(key) {
  return OfficeRuntime.storage.getItem(key);
}

CustomFunctions.associate("STOREVALUE", StoreValue);
CustomFunctions.associate("GETVALUE", GetValue);
```

<span data-ttu-id="3fc0f-113">[Um exemplo de código mais detalhado no GitHub](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Excel-custom-functions/AsyncStorage) fornece um exemplo de passagem destas informações para o painel de tarefas.</span><span class="sxs-lookup"><span data-stu-id="3fc0f-113">[A more detailed code sample on GitHub](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Excel-custom-functions/AsyncStorage) gives an example of passing this information to the task pane.</span></span>

>[!NOTE]
> <span data-ttu-id="3fc0f-114">O objeto `storage` substitui o objeto anterior de armazenamento chamado `AsyncStorage`, que agora se tornou obsoleto.</span><span class="sxs-lookup"><span data-stu-id="3fc0f-114">The `storage` object replaces the previous storage object named `AsyncStorage` which is now deprecated.</span></span> <span data-ttu-id="3fc0f-115">Se o objeto `AsyncStorage` estiver em uso no seu código atual de funções personalizadas, atualize-o para usar o objeto `storage`.</span><span class="sxs-lookup"><span data-stu-id="3fc0f-115">If using the `AsyncStorage` object in your current custom functions code, please update it to use the `storage` object.</span></span>

## <a name="next-steps"></a><span data-ttu-id="3fc0f-116">Próximas etapas</span><span class="sxs-lookup"><span data-stu-id="3fc0f-116">Next steps</span></span>
<span data-ttu-id="3fc0f-117">Saiba como [gerar automaticamente os metadados JSON para as suas funções personalizadas](custom-functions-json-autogeneration.md).</span><span class="sxs-lookup"><span data-stu-id="3fc0f-117">Learn how to [autogenerate the JSON metadata for your custom functions](custom-functions-json-autogeneration.md).</span></span> 

## <a name="see-also"></a><span data-ttu-id="3fc0f-118">Confira também</span><span class="sxs-lookup"><span data-stu-id="3fc0f-118">See also</span></span>

* [<span data-ttu-id="3fc0f-119">Metadados de funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="3fc0f-119">Custom functions metadata</span></span>](custom-functions-json.md)
* [<span data-ttu-id="3fc0f-120">Tempo de execução de funções personalizadas do Excel</span><span class="sxs-lookup"><span data-stu-id="3fc0f-120">Runtime for Excel custom functions</span></span>](custom-functions-runtime.md)
* <span data-ttu-id="3fc0f-121">[Práticas recomendadas de funções personalizadas](custom-functions-best-practices.md).</span><span class="sxs-lookup"><span data-stu-id="3fc0f-121">[Custom functions best practices](custom-functions-best-practices.md)</span></span>
* [<span data-ttu-id="3fc0f-122">Tutorial de funções personalizadas do Excel</span><span class="sxs-lookup"><span data-stu-id="3fc0f-122">Excel custom functions tutorial</span></span>](../tutorials/excel-tutorial-create-custom-functions.md)
* [<span data-ttu-id="3fc0f-123">Depuração de funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="3fc0f-123">Custom functions debugging</span></span>](custom-functions-debugging.md)