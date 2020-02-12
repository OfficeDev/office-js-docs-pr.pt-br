---
ms.date: 07/10/2019
description: Use `OfficeRuntime.storage` para salvar o estado com funções personalizadas.
title: Salvar e compartilhar o estado em funções personalizadas
localization_priority: Normal
ms.openlocfilehash: 397c785a4dedb7d2e9d1b38c8db0edb811448e1d
ms.sourcegitcommit: d15bca2c12732f8599be2ec4b2adc7c254552f52
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 02/12/2020
ms.locfileid: "41950807"
---
# <a name="save-and-share-state-in-custom-functions"></a>Salvar e compartilhar o estado em funções personalizadas

Use o objeto `OfficeRuntime.storage` para salvar o estado relacionado às funções personalizadas ou o painel de tarefas no seu suplemento. O armazenamento é limitado a 10 MB por domínio (que pode ser compartilhado entre vários suplementos). No Excel no Windows, o objeto `storage` é uma localização separada dentro do tempo de execução das funções personalizadas, mas no Excel Online e no Mac, o objeto `storage` é o mesmo que o `localStorage` do navegador.

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

Existem várias maneiras de usar `storage` para o gerenciamento de estado:

- Você pode armazenar valores padrão para funções personalizadas para usar quando você estiver offline e não for possível acessar um recurso da Web.
- Você pode salvar valores para funções personalizadas para evitar fazer chamadas adicionais à um recurso da Web.
- Você pode salvar valores da sua função personalizada.
- Você pode armazenar valores do seu painel de tarefas.

O exemplo de código a seguir ilustra como armazenar um item em `storage` e recuperá-lo.

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
```

[Um exemplo de código mais detalhado no GitHub](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Excel-custom-functions/AsyncStorage) fornece um exemplo de passagem destas informações para o painel de tarefas.

>[!NOTE]
> O objeto `storage` substitui o objeto anterior de armazenamento chamado `AsyncStorage`, que agora se tornou obsoleto. Se o objeto `AsyncStorage` estiver em uso no seu código atual de funções personalizadas, atualize-o para usar o objeto `storage`.

## <a name="next-steps"></a>Próximas etapas
Saiba como [gerar automaticamente os metadados JSON para as suas funções personalizadas](custom-functions-json-autogeneration.md). 

## <a name="see-also"></a>Confira também

* [Metadados de funções personalizadas](custom-functions-json.md)
* [Tempo de execução de funções personalizadas do Excel](custom-functions-runtime.md)
* [Tutorial de funções personalizadas do Excel](../tutorials/excel-tutorial-create-custom-functions.md)
* [Depuração de funções personalizadas](custom-functions-debugging.md)
