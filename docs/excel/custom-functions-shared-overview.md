---
ms.date: 02/13/2020
description: Aprenda a executar funções personalizadas, botões da faixa de opções e código do painel de tarefas no mesmo tempo de execução do JavaScript para coordenar cenários em seu suplemento.
title: Execute seu código de suplemento em um tempo de execução do Javascript compartilhado (visualização)
localization_priority: Priority
ms.openlocfilehash: 774990a9452d450bd5c4d968027bc64ebee858af
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/17/2020
ms.locfileid: "42719529"
---
# <a name="overview-run-your-add-in-code-in-a-shared-javascript-runtime-preview"></a>Visão geral: Execute seu código de suplemento em um tempo de execução do Javascript compartilhado (visualização)

[!include[Running custom functions in shared JavaScript runtime note](../includes/excel-shared-runtime-preview-note.md)]

Ao executar o Excel no Windows ou Mac, o suplemento executará o código para botões da faixa de opções, funções personalizadas e o painel de tarefas em ambientes de tempo de execução JavaScript separados. Isso cria limitações, como não poder compartilhar facilmente dados globais e não poder acessar todas as funcionalidades do CORS a partir de uma função customizada.

No entanto, você pode configurar o suplemento do Excel para compartilhar código no mesmo tempo de execução JavaScript (também conhecido como tempo de execução compartilhado). Isso permite uma melhor coordenação entre o suplemento e o acesso ao DOM e CORS do painel de tarefas de todas as partes do suplemento.

A configuração de um tempo de execução compartilhado permite os seguintes cenários:

- Seu suplemento terá um DOM compartilhado que a faixa de opções, o painel de tarefas e as funções personalizadas podem acessar.
- Suas funções personalizadas terão suporte completo ao CORS.
- Suas funções personalizadas podem chamar as APIs do Office.js para ler os dados do documento da planilha.
- Seu suplemento pode executar o código assim que o documento for aberto.
- Seu suplemento pode continuar executando o código após o fechamento do painel de tarefas.

Quando você executa funções personalizadas em um tempo de execução compartilhado com o painel de tarefas, ele será executado em uma instância do navegador em plataformas diferentes, conforme explicado em [Navegadores usados pelos Suplementos do Office](../concepts/browsers-used-by-office-web-add-ins.md). Além disso, todos os botões exibidos pelo suplemento do Excel na faixa de opções serão executados no mesmo tempo de execução compartilhado. A imagem a seguir mostra como as funções personalizadas, a interface do usuário da faixa de opções e o código do painel de tarefas serão executados no mesmo tempo de execução JavaScript.

![Funções personalizadas em execução no tempo de execução compartilhado com botões da faixa de opções e o painel de tarefas no Excel](../images/custom-functions-in-browser-runtime.png)

## <a name="differences-when-running-custom-functions-in-a-shared-runtime"></a>Diferenças ao executar funções personalizadas em um tempo de execução compartilhado

Quando você configura seu projeto de suplemento do Excel para executar funções personalizadas em um tempo de execução compartilhado, existem algumas diferenças no uso do tempo de execução da função personalizada.

### <a name="storage"></a>Armazenamento

Você não precisa mais usar a API de **Armazenamento** para compartilhar dados entre o painel de tarefas, funções personalizadas ou interface do usuário da faixa de opções. Você pode colocar variáveis globais no objeto de **janela** ou usar sua própria abordagem de gerenciamento de estado preferida.

### <a name="authentication"></a>Autenticação

Quando você recebe tokens como parte da autenticação, não precisa usar a API de **Armazenamento** para compartilhá-los entre o painel de tarefas, funções personalizadas e interface do usuário da faixa de opções. Você pode usar sua própria técnica de armazenamento e local de armazenamento preferidos para compartilhá-los, como `localStorage`.

### <a name="dialog-api"></a>API de Caixa de Diálogo

Você não precisa mais usar a API **OfficeRuntime.Dialog** para exibir uma caixa de diálogo a partir de uma função personalizada. Você pode usar a mesma [API de caixa de diálogo](../develop/dialog-api-in-office-add-ins.md) para funções personalizadas, botões da faixa de opções e o painel de tarefas.

### <a name="debugging"></a>Depuração

Ao usar um tempo de execução compartilhado, não é possível usar o Código do Visual Studio para depurar funções personalizadas no Excel no Windows no momento. Você precisará usar ferramentas de desenvolvedor. Para obter mais informações, consulte [Depurar suplementos usando ferramentas de desenvolvedor no Windows 10](../testing/debug-add-ins-using-f12-developer-tools-on-windows-10.md).

## <a name="get-started"></a>Introdução

Para configurar seu projeto de suplemento do Excel para executar funções personalizadas em um tempo de execução compartilhado, consulte [Configurar o suplemento do Excel para usar um tempo de execução do Javascript compartilhado (visualização)](configure-your-add-in-to-use-a-shared-runtime.md).

## <a name="give-us-feedback"></a>Envie-nos seus comentários

Adoraríamos ouvir seus comentários sobre esse recurso. Se você encontrar algum bug ou problema, ou tiver solicitações sobre esse recurso, informe-nos criando um problema do GitHub no [repositório office-js](https://github.com/OfficeDev/office-js).

## <a name="see-also"></a>Confira também

Lista de artigos relacionados para tempo de execução compartilhado
- [Tutorial: Compartilhar dados e eventos entre as funções personalizadas do Excel e o painel de tarefas (visualização)](../tutorials/share-data-and-events-between-custom-functions-and-the-task-pane-tutorial.md)
- [Chame APIs do Excel a partir de sua função personalizada (visualização)](call-excel-apis-from-custom-function.md)