---
title: Depurar os complementos usando ferramentas de desenvolvedor no Windows
description: Depurar os complementos usando Microsoft Edge ferramentas de desenvolvedor no Windows
ms.date: 10/05/2021
ms.localizationpriority: medium
ms.openlocfilehash: ee3567ffe1b91f49389c78dd7f6df7c4c8d14c2c
ms.sourcegitcommit: 489befc41e543a4fb3c504fd9b3f61322134c1ef
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 10/06/2021
ms.locfileid: "60138804"
---
# <a name="debug-add-ins-using-developer-tools-on-windows"></a>Depurar os complementos usando ferramentas de desenvolvedor no Windows

Há ferramentas de desenvolvedor fora das IDEs disponíveis para ajudá-lo a depurar seus Windows 10 e 11. Elas são úteis quando você precisa investigar um problema enquanto executa seu suplemento fora do IDE.

A ferramenta que você usa depende se o suplemento está sendo executado no Microsoft Edge ou no Internet Explorer. Isso é determinado pela versão do Windows e pela versão Office que estão instaladas no computador. Para determinar qual navegador está sendo usado em seu computador de desenvolvimento, confira [Navegadores usados pelos Suplementos do Office](../concepts/browsers-used-by-office-web-add-ins.md).

> [!NOTE]
> As instruções neste artigo não podem ser utilizadas para depurar um suplemento do Outlook que usa Funções Executar. Para depurar um suplemento do Outlook que usa Funções Executar, é recomendável que você anexe ao Visual Studio no modo de script ou outro depurador de scripts.

## <a name="when-the-add-in-is-running-in-microsoft-edge"></a>Quando o suplemento estiver sendo executado no Microsoft Edge

[!include[Enable debugging on Microsoft Edge DevTools](../includes/enable-debugging-on-edge-devtools.md)]

### <a name="debug-using-microsoft-edge-devtools"></a>Depurar usando o Microsoft Edge DevTools

Quando o suplemento estiver sendo executado, você pode usar o [Microsoft Edge DevTools](https://www.microsoft.com/p/microsoft-edge-devtools-preview/9mzbfrmz0mnj?activetab=pivot%3Aoverviewtab).

1. Execute o suplemento.

2. Execute o Microsoft Edge DevTools.

3. Nas ferramentas, abra a guia **Local**. Seu suplemento será listado por nome.

4. Clique no nome do suplemento para abri-lo nas ferramentas.

5. Abra a guia **Depurador**. 

6. Escolha o ícone de pasta acima do painel **script** (à esquerda). Na lista de arquivos disponíveis exibidos na lista suspensa, selecione o arquivo JavaScript que você deseja depurar.

7. Para definir um ponto de interrupção, selecione a linha. Você verá um ponto vermelho à esquerda da linha e uma linha correspondente no painel **Pilha de Chamadas** (canto inferior direito).

8. Execute funções no suplemento conforme necessário para disparar o ponto de interrupção.

## <a name="when-the-add-in-is-running-in-internet-explorer"></a>Quando o suplemento estiver sendo executado no Internet Explorer

Quando o add-in estiver em execução no Internet Explorer, você poderá usar o depurador das ferramentas de desenvolvedor F12 no Windows para testar o seu complemento. Você pode iniciar as ferramentas de desenvolvedor F12 depois que o suplemento estiver em execução. As ferramentas F12 são exibidas em uma janela separada e não usam o Visual Studio.

> [!NOTE]
> O Depurador faz parte das ferramentas de desenvolvedor F12 Windows 10, 11 e Internet Explorer. As versões anteriores do Windows não incluem o Depurador. 

Este exemplo usa o Word e um suplemento gratuito do AppSource.

1. Abra o Word e escolha um documento em branco. 
    
2. Na guia **Inserir**, no grupo Suplementos e selecione **Store** e selecione o suplemento **QR4Office**. (Você pode carregar qualquer suplemento da Store ou seu catálogo de suplemento.)
    
3. Inicie as ferramentas de desenvolvimento F12 que correspondem à sua versão do Office:
    
   - Para a versão de 32 bits do Office, use C:\Windows\System32\F12\F12Chooser.exe
    
   - Para a versão de 64 bits do Office, use C:\Windows\SysWOW64\F12\F12Chooser.exe
    
   Quando você inicia IEChooser, uma janela separada denominada "Escolher destino para depurar" exibe os possíveis aplicativos para depurar. Selecione o aplicativo do seu interesse. Se você estiver escrevendo seu próprio suplemento, selecione o site onde você deseja ter o suplemento implantado, que pode ser uma URL de localhost. 
    
   Por exemplo, selecione **home.html**. 
    
   ![Tela IEChooser, apontando para o complemento bolhas.](../images/choose-target-to-debug.png)

4. Na janela F12, selecione o arquivo que você deseja depurar.
    
   Para selecionar o arquivo na janela F12, escolha o ícone de pasta acima do painel **script** (à esquerda). Na lista de arquivos disponíveis exibido na lista suspensa, selecione **Home.js**.
    
5. Defina o ponto de interrupção.
    
   Para definir o ponto de interrupção no **Home.js**, escolha a linha 144, que está na função  `textChanged`. Você verá um ponto vermelho à esquerda da linha e uma linha correspondente no painel Pilha de Chamadas e Pontos de Interrupção (canto inferior direito). Para ver outras maneiras de definir um ponto de interrupção, confira [Inspecionar executando JavaScript com o Depurador](/previous-versions/windows/internet-explorer/ie-developer/samples/dn255007(v=vs.85)). 
    
   ![Depurador com ponto de interrupção home.js arquivo.](../images/debugger-home-js-02.png)

6. Execute o suplemento para acionar o ponto de interrupção.
    
   No Word, escolha a caixa de texto na parte superior da URL do painel **QR4Office** e tente digitar algum texto. No Depurador, no painel **Pilha de Chamadas e Pontos de Interrupção**, você verá que o ponto de interrupção está disparado e mostra várias informações. Talvez você precise atualizar o depurador para ver os resultados.
    
   ![Depurador com resultados do ponto de interrupção disparado.](../images/debugger-home-js-01.png)


## <a name="see-also"></a>Confira também

- [Inspecionar executando JavaScript com o Depurador](/previous-versions/windows/internet-explorer/ie-developer/samples/dn255007(v=vs.85))
- [Usando as ferramentas de desenvolvedor F12](/previous-versions/windows/internet-explorer/ie-developer/samples/bg182326(v=vs.85))
