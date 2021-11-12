---
title: Depurar os complementos usando ferramentas de desenvolvedor para Microsoft Edge WebView2
description: Depurar os complementos usando as ferramentas de desenvolvedor no Microsoft Edge WebView2 (Chromium baseados em Chromium).
ms.date: 11/09/2021
ms.localizationpriority: medium
ms.openlocfilehash: d2a8aa41720eebcd15d4cb625d90af1399b87dad
ms.sourcegitcommit: 3d37c42f5e465dac52d231d31717bdbb3bfa0e30
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 11/10/2021
ms.locfileid: "60923535"
---
# <a name="debug-add-ins-using-developer-tools-in-microsoft-edge-chromium-based"></a>Depurar os complementos usando ferramentas de desenvolvedor no Microsoft Edge (Chromium baseados em Chromium)

Este artigo mostra como depurar o código do lado do cliente (JavaScript ou TypeScript) do seu complemento quando as seguintes condições são atendidas.

- Você não pode ou não deseja depurar usando ferramentas criadas em seu IDE; ou você está encontrando um problema que só ocorre quando o complemento é executado fora do IDE.
- Seu computador está usando uma combinação de Windows e Office que usam o controle webview Edge (Chromium baseado em Chromium), WebView2.

> [!TIP]
> Para obter informações sobre a depuração com o Edge WebView2 (baseado em Chromium) no Visual Studio Code, consulte Depurar os complementos no Windows usando Visual Studio Code e [Microsoft Edge WebView2 (baseados em Chromium)](debug-desktop-using-edge-chromium.md).

Para determinar qual navegador você está usando, consulte [Browsers used by Office Add-ins](../concepts/browsers-used-by-office-web-add-ins.md).

> [!TIP]
> [!INCLUDE[Identify the webview through the add-in UI](../includes/identify-webview-in-ui.md)]

## <a name="debug-a-task-pane-add-in-using-microsoft-edge-chromium-based-developer-tools"></a>Depurar um add-in do painel de tarefas usando Microsoft Edge ferramentas de desenvolvedor (Chromium baseadas em Chromium)

> [!NOTE]
> Se o seu add-in tiver um comando de complemento que execute uma função, a função será executada em um processo de navegador oculto do qual as ferramentas de desenvolvedor do Microsoft Edge (baseadas em Chromium) não podem ser lançadas, portanto, a técnica descrita neste artigo não pode ser usada para depurar o código na função. [](../design/add-in-commands.md)

1. [Fazer sideload](create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md) e executar o complemento.
1. Execute as Microsoft Edge de desenvolvedor (Chromium baseadas em Chromium) por um destes métodos:

   - Certifique-se de que o painel de tarefas do complemento tenha foco e pressione **Ctrl+Shift+I**.
   - Clique com o botão direito do mouse no painel de tarefas para abrir o menu de contexto e selecione **Inspecionar**, ou abra o [menu](../design/task-pane-add-ins.md#personality-menu) personalidade e selecione **Anexar Depurador**.

1. Abra a **guia Fontes.**
1. Abra o arquivo que você deseja depurar com as etapas a seguir.

   1. À direita da barra de menus superior da ferramenta, selecione o **botão ...** e selecione **Pesquisar**.
   1. Insira uma linha de código do arquivo que você deseja depurar na caixa de pesquisa. Deve ser algo que provavelmente não estará em nenhum outro arquivo.
   1. Selecione o botão atualizar.
   1. Nos resultados da pesquisa, selecione a linha para abrir o arquivo de código no painel acima dos resultados da pesquisa.

   :::image type="content" source="../images/open-file-in-edge-chromium-devtools.png" alt-text="Captura de tela da guia de origem Chromium ferramentas de desenvolvedor de borda com 4 partes rotuladas de A a D.":::

1. Para definir um ponto de interrupção, selecione o número de linha da linha no arquivo de código. Um ponto vermelho é exibido pela linha no arquivo de código. Na janela de depurador à direita, o ponto de interrupção é registrado na lista de pontos **de** interrupção.
1. Execute funções no suplemento conforme necessário para disparar o ponto de interrupção.

> [!TIP]
> Para obter mais informações sobre como usar as ferramentas, consulte Microsoft Edge visão geral das Ferramentas de [Desenvolvedor.](/microsoft-edge/devtools-guide-chromium/)

## <a name="debug-a-dialog-in-an-add-in"></a>Depurar uma caixa de diálogo em um complemento

Se o seu add-in usar a API de diálogo Office, a caixa de diálogo será executado em um processo separado do painel de tarefas (se algum) e a ferramenta deverá ser iniciada a partir desse processo separado. Siga estas etapas.

1. Execute o suplemento.
1. Abra a caixa de diálogo e certifique-se de que ela tenha foco.
1. Abra as Microsoft Edge de desenvolvedor (Chromium baseadas em Chromium) por um destes métodos:

   - Pressione **Ctrl+Shift+I** ou **F12**.
   - Clique com o botão direito do mouse na caixa de diálogo para abrir o menu de contexto e selecione **Inspecionar**.

1. Use a ferramenta da mesma forma que você faria para código em um painel de tarefas. Consulte [Depurar um add-in](#debug-a-task-pane-add-in-using-microsoft-edge-chromium-based-developer-tools) do painel de tarefas usando Microsoft Edge ferramentas de desenvolvedor (Chromium baseadas em Chromium) anteriormente neste artigo.
