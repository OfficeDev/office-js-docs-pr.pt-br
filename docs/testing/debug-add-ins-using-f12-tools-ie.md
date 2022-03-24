---
title: Depurar os suplementos usando as ferramentas de desenvolvedor para o Internet Explorer
description: Depurar os complementos usando as ferramentas de desenvolvedor no Internet Explorer.
ms.date: 11/02/2021
ms.localizationpriority: medium
ms.openlocfilehash: bb7c328e6b1f839d5d711f81beceaf7519545fe3
ms.sourcegitcommit: 968d637defe816449a797aefd930872229214898
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/23/2022
ms.locfileid: "63744671"
---
# <a name="debug-add-ins-using-developer-tools-in-internet-explorer"></a>Depurar os complementos usando ferramentas de desenvolvedor no Internet Explorer

Este artigo mostra como depurar o código do lado do cliente (JavaScript ou TypeScript) do seu complemento quando as seguintes condições são atendidas.

- Você não pode ou não deseja depurar usando ferramentas criadas em seu IDE; ou você está encontrando um problema que só ocorre quando o complemento é executado fora do IDE.
- Seu computador está usando uma combinação de Windows e Office que usam o controle webview do Internet Explorer, Trident.

Para determinar qual navegador está sendo usado em seu computador, consulte [Browsers used by Office Add-ins](../concepts/browsers-used-by-office-web-add-ins.md).

> [!TIP]
> [!INCLUDE[Identify the webview through the add-in UI](../includes/identify-webview-in-ui.md)]

> [!NOTE]
> Para instalar uma versão do Office que usa o Webview do Internet Explorer ou para forçar sua versão atual a usar o Internet Explorer, consulte [Switch to the Internet Explorer 11 webview](#switch-to-the-internet-explorer-11-webview).

## <a name="debug-a-task-pane-add-in-using-the-f12-tools"></a>Depurar um complemento do painel de tarefas usando as ferramentas F12

Windows 10 e 11 incluem uma ferramenta de desenvolvimento da Web chamada "F12" porque foi originalmente lançada pressionando F12 no Internet Explorer. O F12 agora é um aplicativo independente usado para depurar seu complemento quando ele está sendo executado no controle webview do Internet Explorer, Trident. O aplicativo não está disponível em versões anteriores do Windows.

> [!NOTE]
> Se o seu add-in tiver [](../design/add-in-commands.md) um comando de complemento que execute uma função, a função será executada em um processo de navegador oculto ao qual as ferramentas F12 não podem detectar ou anexar, portanto, a técnica descrita neste artigo não pode ser usada para depurar o código na função.

As etapas a seguir são as instruções para depurar seu complemento. Se você quiser testar as próprias ferramentas F12, consulte [Exemplo de complemento para testar as ferramentas F12](#example-add-in-to-test-the-f12-tools).

1. [Fazer sideload](create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md) e executar o complemento.
1. Iniciar as ferramentas de desenvolvimento F12 que correspondem à sua versão do Office.

   - Para a versão de 32 bits do Office, use C:\Windows\System32\F12\F12Chooser.exe
   - Para a versão de 64 bits do Office, use C:\Windows\SysWOW64\F12\F12Chooser.exe

   O IEChooser é aberto com uma janela chamada **Escolher destino a ser depurado**. Seu complemento aparecerá na janela nomeada pelo nome do arquivo da home page do complemento. Na captura de tela a seguir, é `Home.html`. Apenas os processos que estão sendo executados no Internet Explorer ou no Trident são exibidos. A ferramenta não pode ser anexada a processos que estão sendo executados em outros navegadores ou webviews, incluindo Microsoft Edge.

    :::image type="content" source="../images/choose-target-to-debug.png" alt-text="Tela IEChooser, com vários processos internet Explorer e Trident listados. Um é chamado Home.html.":::

1. Selecione o processo do seu complemento; ou seja, seu nome de arquivo de página inicial. Essa ação anexa as ferramentas F12 ao processo e abre a interface principal do usuário F12.
1. Abra a guia **Depurador**.
1. No canto superior esquerdo da guia, logo abaixo da faixa de opções da ferramenta de depurador, há um pequeno ícone de pasta. Selecione isso para abrir uma listada dos arquivos no complemento. Apresentamos um exemplo a seguir.

    :::image type="content" source="../images/f12-file-dropdown.png" alt-text="Captura de tela do canto superior esquerdo da guia depurador com uma pasta listada aberta e uma lista de arquivos.":::

1. Selecione o arquivo que você deseja depurar e ele será aberto no **painel de script** (à esquerda) da guia **Depurador** . Se você estiver usando um transpiler, empacotador ou minifier, que altera o nome do arquivo, ele terá o nome final que é realmente carregado, não o nome do arquivo de origem original.

1. Role para uma linha onde você deseja definir um ponto de interrupção e clique na margem à esquerda do número da linha. Você verá um ponto vermelho à esquerda da linha e uma linha correspondente aparece na guia **Pontos** de Interrupção do painel inferior direito. A captura de tela a seguir é um exemplo.

    :::image type="content" source="../images/debugger-home-js-02.png" alt-text="Depurador com ponto de interrupção home.js arquivo.":::

1. Execute funções no suplemento conforme necessário para disparar o ponto de interrupção. Quando o ponto de interrupção é atingido, uma seta apontando para a direita aparece no ponto vermelho do ponto de interrupção. A captura de tela a seguir é um exemplo.

    :::image type="content" source="../images/debugger-home-js-01.png" alt-text="Depurador com resultados do ponto de interrupção disparado.":::

> [!TIP]
> Para obter mais informações sobre como usar as ferramentas F12, consulte [Inspect running JavaScript with the Debugger](/previous-versions/windows/internet-explorer/ie-developer/samples/dn255007(v=vs.85)).

### <a name="example-add-in-to-test-the-f12-tools"></a>Exemplo de complemento para testar as ferramentas F12

Este exemplo usa o Word e um suplemento gratuito do AppSource.

1. Abra o Word e escolha um documento em branco.
1. Na guia **Inserir**, no grupo **Add-ins**, selecione **Meus Complementos** para abrir **a** caixa de diálogo Office de Office e selecione a **guia STORE**.
1. Selecione o **complemento QR4Office** . Ele abre em um painel de tarefas.
1. Iniciar as ferramentas de desenvolvimento F12 que correspondem à sua versão do Office conforme descrito na seção anterior.
1. Na janela F12, selecione **Home.html**.
1. Na guia **Depurador** , abra o **arquivoHome.jsconforme** descrito na seção anterior.
1. De definir os pontos de interrupção nas linhas 310 e 312.
1. No complemento, selecione o **botão Inserir** . Um ou outro ponto de interrupção é atingido.

## <a name="debug-a-dialog-in-an-add-in"></a>Depurar uma caixa de diálogo em um complemento

Se o seu add-in usar a API de caixa de diálogo Office, a caixa de diálogo será executado em um processo separado do painel de tarefas (se algum) e as ferramentas deverão ser anexados a esse processo. Siga estas etapas.

1. Execute o complemento e as ferramentas. 
1. Abra a caixa de diálogo e selecione o **botão Atualizar** nas ferramentas. O processo de caixa de diálogo é mostrado. Seu nome é o nome do arquivo que está aberto na caixa de diálogo.
1. Selecione o processo para abri-lo e depurar conforme descrito na seção Depurar um complemento do painel de [tarefas usando as ferramentas F12](#debug-a-task-pane-add-in-using-the-f12-tools).

## <a name="switch-to-the-internet-explorer-11-webview"></a>Alternar para o Webview do Internet Explorer 11

Há duas maneiras de alternar o Webview do Internet Explorer. Você pode executar um comando simples em um prompt de comando ou instalar uma versão do Office que usa o Internet Explorer por padrão. Recomendamos o primeiro método. Mas você deve usar o segundo nos cenários a seguir.

- Seu projeto foi desenvolvido com Visual Studio e IIS. Não é baseado em node.js.
- Você deseja ser absolutamente robusto em seus testes.
- Se, por qualquer motivo, a ferramenta de linha de comando não funcionar.

### <a name="switch-via-the-command-line"></a>Alternar pela linha de comando

[!INCLUDE [Steps to switch browsers with the command line tool](../includes/use-legacy-edge-or-ie.md)]

### <a name="install-a-version-of-office-that-uses-internet-explorer"></a>Instalar uma versão do Office que usa o Internet Explorer

[!INCLUDE [Steps to install Office that uses Edge Legacy or Internet Explorer](../includes/install-office-that-uses-legacy-edge-or-ie.md)]

## <a name="see-also"></a>Confira também

- [Inspecionar executando JavaScript com o Depurador](/previous-versions/windows/internet-explorer/ie-developer/samples/dn255007(v=vs.85))
- [Usando as ferramentas de desenvolvedor F12](/previous-versions/windows/internet-explorer/ie-developer/samples/bg182326(v=vs.85))
