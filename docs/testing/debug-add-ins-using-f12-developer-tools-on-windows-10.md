---
title: Depurar suplementos usando as ferramentas de desenvolvedor F12 no Windows 10
description: ''
ms.date: 03/19/2019
localization_priority: Priority
ms.openlocfilehash: 750411bea187a0ade9b3723e3198d82f7c482c9f
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 04/24/2019
ms.locfileid: "32450135"
---
# <a name="debug-add-ins-using-f12-developer-tools-on-windows-10"></a>Depurar suplementos usando as ferramentas de desenvolvedor F12 no Windows 10

As ferramentas de desenvolvedor F12 incluídas no Windows 10 o ajudam a depurar, testar e acelerar suas páginas da Web. Você também pode usá-las para desenvolver e a depurar seu suplemento do Office se não estiver usando um IDE como o Visual Studio ou se precisar investigar um problema durante a execução do suplemento fora do IDE. Este artigo mostra como é possível usar a ferramenta Depurador das ferramentas de desenvolvedor F12 no Windows 10 para testar seu suplemento do Office.

> [!NOTE]
> As instruções neste artigo não podem ser utilizadas para depurar um suplemento do Outlook que usa Funções Executar. Para depurar um suplemento do Outlook que usa Funções Executar, é recomendável que você anexe ao Visual Studio no modo de script ou outro depurador de scripts.

## <a name="prerequisites"></a>Pré-requisitos

Você precisa dos seguintes softwares:

- As ferramentas do desenvolvedor F12, que estão incluídas no Windows 10. 
    
- O aplicativo cliente do Office que hospeda seu suplemento. 
    
- Seu suplemento. 

## <a name="using-the-debugger"></a>Utilização do depurador

Este artigo mostra como é possível usar a ferramenta Depurador das ferramentas de desenvolvedor F12 no Windows 10 para testar seu suplemento do Office. É possível testar os suplementos do AppSource ou os suplementos que você adicionou de outros locais. As ferramentas F12 são exibidas em uma janela separada e não usam o Visual Studio. Você pode iniciar as ferramentas de desenvolvedor F12 depois que o suplemento estiver em execução. As ferramentas F12 são exibidas em uma janela separada e não usam o Visual Studio.

> [!NOTE]
> O Depurador faz parte das ferramentas de desenvolvedor F12 no Windows 10 e no Internet Explorer. As versões anteriores do Windows não incluem o Depurador. 

Este exemplo usa o Word e um suplemento gratuito do AppSource.

1. Abra o Word e escolha um documento em branco. 
    
2. Na guia **Inserir**, no grupo Suplementos e selecione **Store** e selecione o suplemento **QR4Office**. (Você pode carregar qualquer suplemento da Store ou seu catálogo de suplemento.)
    
3. Inicie as ferramentas de desenvolvimento F12 que correspondem à sua versão do Office:
    
   - Para a versão de 32 bits do Office, use C:\Windows\System32\F12\F12Chooser.exe
    
   - Para a versão de 64 bits do Office, use C:\Windows\SysWOW64\F12\F12Chooser.exe
    
   Quando você inicia IEChooser, uma janela separada denominada "Escolher destino para depurar" exibe os possíveis aplicativos para depurar. Selecione o aplicativo do seu interesse. Se você estiver escrevendo seu próprio suplemento, selecione o site onde você deseja ter o suplemento implantado, que pode ser uma URL de localhost. 
    
   Por exemplo, selecione **home.html**. 
    
   ![Tela do IEChooser, apontando para o suplemento bolhas](../images/choose-target-to-debug.png)

4. Na janela F12, selecione o arquivo que você deseja depurar.
    
   Para selecionar o arquivo na janela F12, escolha o ícone de pasta acima do painel **script** (à esquerda). Na lista de arquivos disponíveis exibido na lista suspensa, selecione **Home.js**.
    
5. Defina o ponto de interrupção.
    
   Para definir o ponto de interrupção no **Home.js**, escolha a linha 144, que está na função  `textChanged`. Você verá um ponto vermelho à esquerda da linha e uma linha correspondente no painel Pilha de Chamadas e Pontos de Interrupção (canto inferior direito). Para ver outras maneiras de definir um ponto de interrupção, confira [Inspecionar executando JavaScript com o Depurador](/previous-versions/windows/internet-explorer/ie-developer/samples/dn255007(v=vs.85)). 
    
   ![Depurador com ponto de interrupção no arquivo home.js](../images/debugger-home-js-02.png)

6. Execute o suplemento para acionar o ponto de interrupção.
    
   No Word, escolha a caixa de texto na parte superior da URL do painel **QR4Office** e tente digitar algum texto. No Depurador, no painel **Pilha de Chamadas e Pontos de Interrupção**, você verá que o ponto de interrupção está disparado e mostra várias informações. Talvez você precise atualizar o depurador para ver os resultados.
    
   ![Depurador com resultados do ponto de interrupção disparado](../images/debugger-home-js-01.png)


## <a name="see-also"></a>Confira também

- [Inspecionar executando JavaScript com o Depurador](/previous-versions/windows/internet-explorer/ie-developer/samples/dn255007(v=vs.85))
- [Usando as ferramentas de desenvolvedor F12](/previous-versions/windows/internet-explorer/ie-developer/samples/bg182326(v=vs.85))
