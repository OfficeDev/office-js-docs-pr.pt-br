---
ms.date: 05/17/2020
description: Aprenda a executar funções personalizadas, botões da faixa de opções e código do painel de tarefas no mesmo tempo de execução do JavaScript para coordenar cenários em seu suplemento.
title: Executar o código do suplemento em um tempo de execução de JavaScript compartilhado
localization_priority: Priority
ms.openlocfilehash: e179d8cdff06b397c9998a6a18a506a0a6df1b42
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 06/08/2020
ms.locfileid: "44609223"
---
# <a name="overview-run-your-add-in-code-in-a-shared-javascript-runtimes"></a>Visão geral: executar o código do suplemento em um ambiente de tempo de execução JavaScript compartilhado

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

Ao executar o Excel no Windows ou Mac, o suplemento executará o código para botões da faixa de opções, funções personalizadas e o painel de tarefas em ambientes de tempo de execução JavaScript separados. Isso cria limitações, como não poder compartilhar facilmente dados globais e não poder acessar todas as funcionalidades do CORS a partir de uma função customizada.

No entanto, você pode configurar o suplemento do Excel para compartilhar código no mesmo tempo de execução JavaScript (também conhecido como tempo de execução compartilhado). Isso permite uma melhor coordenação entre o suplemento e o acesso ao DOM e CORS do painel de tarefas de todas as partes do suplemento.

A configuração de um tempo de execução compartilhado permite os seguintes cenários:

- Seu suplemento terá um DOM compartilhado que a faixa de opções, o painel de tarefas e as funções personalizadas podem acessar.
- Suas funções personalizadas terão suporte completo ao CORS.
- Suas funções personalizadas podem chamar as APIs do Office.js para ler os dados do documento da planilha.
- Seu suplemento pode executar o código assim que o documento for aberto.
- Seu suplemento pode continuar executando o código após o fechamento do painel de tarefas.

Quando você executa funções personalizadas em um tempo de execução compartilhado com o painel de tarefas, ele será executado em uma instância do navegador em plataformas diferentes, conforme explicado em [Navegadores usados pelos Suplementos do Office](../concepts/browsers-used-by-office-web-add-ins.md). Além disso, todos os botões exibidos pelo suplemento do Excel na faixa de opções serão executados no mesmo tempo de execução compartilhado. A imagem a seguir mostra como as funções personalizadas, a interface do usuário da faixa de opções e o código do painel de tarefas serão executados no mesmo tempo de execução JavaScript.

![Funções personalizadas em execução em um tempo de execução compartilhado com botões de faixa de opções e o painel de tarefas no Excel](../images/custom-functions-in-browser-runtime.png)

## <a name="set-up-a-shared-runtime"></a>Configurar um tempo de execução compartilhado

Consulte o [artigo Configurando um tempo de execução compartilhado](./configure-your-add-in-to-use-a-shared-runtime.md) para saber como configurar suas funções personalizadas para usar um tempo de execução compartilhado.

### <a name="debugging"></a>Depuração

Ao usar um tempo de execução compartilhado, não é possível usar o Código do Visual Studio para depurar funções personalizadas no Excel no Windows no momento. Em vez disso, você precisará usar as ferramentas de desenvolvedor. Para obter mais informações, consulte [Depurar suplementos usando ferramentas de desenvolvedor no Windows 10](../testing/debug-add-ins-using-f12-developer-tools-on-windows-10.md).

## <a name="give-us-feedback"></a>Envie-nos seus comentários

Adoraríamos ouvir seus comentários sobre esse recurso. Se você encontrar algum bug ou problema, ou tiver solicitações sobre esse recurso, informe-nos criando um problema do GitHub no [repositório office-js](https://github.com/OfficeDev/office-js).

## <a name="see-also"></a>Confira também

- [Tutorial: compartilhar dados e eventos entre as funções personalizadas do Excel e o painel de tarefas](../tutorials/share-data-and-events-between-custom-functions-and-the-task-pane-tutorial.md)
- [Chamar as APIs do Excel a partir de sua função personalizada](call-excel-apis-from-custom-function.md)
