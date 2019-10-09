---
ms.date: 07/10/2019
description: Saiba mais sobre o tempo de execução de funções personalizadas do Excel.
title: Arquitetura de funções personalizadas
localization_priority: Priority
ms.openlocfilehash: abe4f847069b3bb9d3813b4520bf8eb078a40c18
ms.sourcegitcommit: bb44c9694f88cde32ffbb642689130db44456964
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 07/17/2019
ms.locfileid: "35771460"
---
# <a name="custom-functions-architecture"></a>Arquitetura de funções personalizadas

 Funções personalizadas estão com tempos de execução exclusivos que priorizam a execução de cálculos. Este artigo aborda a diferenças entre o tempo de execução de funções personalizadas e o mecanismo de JavaScript baseados em navegador que habilita a maioria das outras partes do suplemento.

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

## <a name="custom-functions-runtime"></a>Tempo de execução de funções personalizadas

Um suplemento Web do Office pode interagir com o usuário como um painel de tarefas ou como um painel de conteúdo e pode incluir comandos e funções personalizadas. Todas essas partes são executadas em um tempo de execução do mecanismo do navegador, exceto para funções personalizadas. As funções personalizadas são executadas em um tempo de execução de funções personalizadas separado para otimizar a velocidade de cálculo.

Observe que, se você estiver usando o [gerador Yeoman de suplementos do Office](https://www.npmjs.com/package/generator-office) para gerar o seu projeto, o tempo de execução de funções personalizadas será carregado por meio do arquivo de script custom-functions.js mencionado no arquivo **functions.html**. O **functions.html** serve apenas para carregar o tempo de execução e não deve ser usado como painel de tarefas para o suplemento.

A tabela a seguir destaca as diferenças entre o tempo de execução de funções personalizadas e o tempo de execução do mecanismo do navegador:

| Tempo de execução de funções personalizadas  | Tempo de execução do mecanismo do navegador    |
|------------------------------------------------------------------ |-------------------------------------------------------------------------------------------------------------- |
| Suporte para retornar o valor de uma célula    | Suporte para APIs Office.js e elementos de interface do usuário   |
| Não há o objeto `localStorage`, em vez disso, usa-se o objeto `OfficeRuntime.storage`.     | Há o objeto `localStorage`, opcionalmente poderá usar o objeto`OfficeRuntime.storage`.     |
| Não suporta a interação com o DOM ou o carregamento de  bibliotecas que dependem do DOM, como jQuery.    | Suporta interação com o DOM e com o carregamento de bibliotecas que dependem do DOM. |

## <a name="browser-engine-runtime"></a>Tempo de execução do mecanismo do navegador

O painel de tarefas o, suplemento de conteúdo e os comandos são executados em um tempo de execução do mecanismo do navegador.

APIs Office.js é compatível com o tempo de execução do mecanismo do navegador. Tenha em mente que qualquer uma das APIs do Excel, assim como APIs que permitem manipular tabelas do Excel, são executadas no tempo de execução do mecanismo do navegador, mas não são acessíveis diretamente do tempo de execução de funções personalizadas.

## <a name="communicate-between-runtimes"></a>Comunicar-se entre os tempos de execução

O código de funções personalizadas não pode interagir diretamente com o código em outras partes do seu suplemento da web, como o painel de tarefas porque estão em diferentes tempos de execução. Mas em alguns cenários, talvez seja necessário compartilhar dados, por exemplo, passando um token.

O `OfficeRuntime.storage` pode ser usado para armazenar dados de suas funções personalizadas e para obter dados do seu código do painel de tarefas. Para saber mais sobre como armazenar e compartilhar dados, confira [Salvar e compartilhar estado](custom-functions-save-state.md).

Você pode ver um exemplo de código usando o objeto `storage` neste [repositório Github](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Excel-custom-functions/AsyncStorage) dedicado para padrões e práticas recomendadas.
Para saber mais sobre o objeto `storage`, confira [Tempo de execução de funções personalizadas](./custom-functions-runtime.md).

O objeto `storage` também pode ser útil para autenticação. Para saber mais, confira [autenticação de funções personalizadas](custom-functions-authentication.md).

## <a name="next-steps"></a>Próximas etapas
Saiba mais sobre como [usar o tempo de execução de funções personalizadas](custom-functions-runtime.md)..

## <a name="see-also"></a>Confira também

* [Tempo de execução de funções personalizadas do Excel](custom-functions-runtime.md)
* [Receber e tratar dados com funções personalizadas](custom-functions-web-reqs.md)
* [Metadados de funções personalizadas](custom-functions-json.md)
* [Tutorial de funções personalizadas do Excel](../tutorials/excel-tutorial-create-custom-functions.md)
