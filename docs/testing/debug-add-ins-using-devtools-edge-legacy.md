---
title: Depurar os complementos usando ferramentas de desenvolvedor para Versão Prévia do Microsoft Edge
description: Depurar os complementos usando as ferramentas de desenvolvedor no Versão Prévia do Microsoft Edge.
ms.date: 11/02/2021
ms.localizationpriority: medium
ms.openlocfilehash: 62f27e2ee266e3b6a92d090e8008b74bac4a9663
ms.sourcegitcommit: 968d637defe816449a797aefd930872229214898
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/23/2022
ms.locfileid: "63744679"
---
# <a name="debug-add-ins-using-developer-tools-in-microsoft-edge-legacy"></a>Depurar os complementos usando ferramentas de desenvolvedor no Versão Prévia do Microsoft Edge

Este artigo mostra como depurar o código do lado do cliente (JavaScript ou TypeScript) do seu complemento quando as seguintes condições são atendidas.

- Você não pode ou não deseja depurar usando ferramentas criadas em seu IDE; ou você está encontrando um problema que só ocorre quando o complemento é executado fora do IDE.
- Seu computador está usando uma combinação de Windows e Office que usam o controle webview de Borda original, EdgeHTML.

> [!TIP]
> Para obter informações sobre a depuração com o Legado de Borda dentro Visual Studio Code, consulte [Microsoft Office Extensão de Depurador de Visual Studio Code](debug-with-vs-extension.md).

Para determinar qual navegador você está usando, consulte [Browsers used by Office Add-ins](../concepts/browsers-used-by-office-web-add-ins.md). 

> [!TIP]
> [!INCLUDE[Identify the webview through the add-in UI](../includes/identify-webview-in-ui.md)]

> [!NOTE]
> Para instalar uma versão do Office que usa a webview herdada de Borda ou para forçar sua versão atual do Office a usar o Edge Legacy, consulte [Switch to the Edge Legacy webview](#switch-to-the-edge-legacy-webview).

## <a name="debug-a-task-pane-add-in-using-microsoft-edge-devtools-preview"></a>Depurar um complemento do painel de tarefas usando Microsoft Edge Visualização do DevTools

1. Instale o [Microsoft Edge Visualização do DevTools](https://www.microsoft.com/p/microsoft-edge-devtools-preview/9mzbfrmz0mnj?activetab=pivot%3Aoverviewtab). (A palavra "Visualização" está no nome por motivos históricos. Não há uma versão mais recente.)

   > [!NOTE]
   > Se o seu add-in tiver [](../design/add-in-commands.md) um comando de complemento que execute uma função, a função será executada em um processo de navegador oculto ao qual o Microsoft Edge DevTools não pode detectar ou anexar, portanto, a técnica descrita neste artigo não pode ser usada para depurar o código na função.

1. [Fazer sideload](create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md) e executar o complemento.
1. Execute o Microsoft Edge DevTools.
1. Nas ferramentas, abra a guia **Local**. Seu suplemento será listado por nome. (Somente os processos que estão sendo executados no EdgeHTML aparecem na guia. A ferramenta não pode ser anexada a processos que estão sendo executados em outros navegadores ou webviews, incluindo Microsoft Edge (WebView2) e Internet Explorer (Trident).)

   :::image type="content" source="../images/edge-devtools-with-add-in-process.png" alt-text="Captura de tela do Edge DevTools mostrando um processo chamado legacy-edge-debugging.":::

1. Selecione o nome do complemento para abri-lo nas ferramentas.
1. Abra a guia **Depurador**.
1. Abra o arquivo que você deseja depurar com as etapas a seguir.

   1. Na barra de tarefas de depurador, selecione **Mostrar encontrar em arquivos**. Isso abrirá uma janela de pesquisa.
   1. Insira uma linha de código do arquivo que você deseja depurar na caixa de pesquisa. Deve ser algo que provavelmente não estará em nenhum outro arquivo.
   1. Selecione o botão atualizar.
   1. Nos resultados da pesquisa, selecione a linha para abrir o arquivo de código no painel acima dos resultados da pesquisa.

   :::image type="content" source="../images/open-file-in-edge-devtools.png" alt-text="Captura de tela da guia de depuração de Edge DevTools com 4 partes rotuladas de A a D.":::

1. Para definir um ponto de interrupção, selecione a linha no arquivo de código. O ponto de interrupção é registrado no painel **Pilha de chamada** (inferior direito). Também pode haver um ponto vermelho pela linha no arquivo de código, mas isso não parece confiável.
1. Execute funções no suplemento conforme necessário para disparar o ponto de interrupção.

> [!TIP]
> Para obter mais informações sobre como usar as ferramentas, [consulte Microsoft Edge (EdgeHTML) Developer Tools](/archive/microsoft-edge/legacy/developer/devtools-guide/).

## <a name="debug-a-dialog-in-an-add-in"></a>Depurar uma caixa de diálogo em um complemento

Se o seu add-in usar a API de caixa de diálogo Office, a caixa de diálogo será executado em um processo separado do painel de tarefas (se algum) e as ferramentas deverão ser anexados a esse processo. Siga estas etapas.

1. Execute o complemento e as ferramentas.
1. Abra a caixa de diálogo e selecione o **botão Atualizar** nas ferramentas. O processo de caixa de diálogo é mostrado. Seu nome vem do elemento `<title>` no arquivo HTML que está aberto na caixa de diálogo.
1. Selecione o processo para abri-lo e depurar conforme descrito na seção Depurar um complemento do painel de tarefas usando Microsoft Edge [Visualização do DevTools](#debug-a-task-pane-add-in-using-microsoft-edge-devtools-preview).

   :::image type="content" source="../images/edge-devtools-with-add-in-and-dialog-processes.png" alt-text="Captura de tela de Edge DevTools mostrando um processo chamado Minha Caixa de Diálogo.":::

## <a name="switch-to-the-edge-legacy-webview"></a>Alternar para a webview herdda de borda

Há duas maneiras de alternar o modo webview herddo de borda. Você pode executar um comando simples em um prompt de comando ou pode instalar uma versão do Office que usa Edge Legacy por padrão. Recomendamos o primeiro método. Mas você deve usar o segundo nos cenários a seguir.

- Seu projeto foi desenvolvido com Visual Studio e IIS. Não é baseado em node.js.
- Você deseja ser absolutamente robusto em seus testes.
- Se, por qualquer motivo, a ferramenta de linha de comando não funcionar.

### <a name="switch-via-the-command-line"></a>Alternar pela linha de comando

[!INCLUDE [Steps to switch browsers with the command line tool](../includes/use-legacy-edge-or-ie.md)]

### <a name="install-a-version-of-office-that-uses-edge-legacy"></a>Instalar uma versão do Office que usa o Edge Legacy

[!INCLUDE [Steps to install Office that uses Edge Legacy or Internet Explorer](../includes/install-office-that-uses-legacy-edge-or-ie.md)]
