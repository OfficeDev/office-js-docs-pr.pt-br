---
title: Depurar suplementos usando as ferramentas de desenvolvedor F12 no Windows 10
description: ''
ms.date: 10/16/2018
ms.openlocfilehash: 3df245fcd651ec227e0a32d53da186ee332beb8f
ms.sourcegitcommit: f47654582acbe9f618bec49fb97e1d30f8701b62
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 10/17/2018
ms.locfileid: "25579839"
---
# <a name="debug-add-ins-using-f12-developer-tools-on-windows-10"></a>Depurar suplementos usando as ferramentas de desenvolvedor F12 no Windows 10

As ferramentas de desenvolvedor F12 incluídas no Windows 10 irão ajudá-lo a depurar, testar e acelerar suas páginas da Web. Você também pode usá-las para desenvolver e depurar suplementos do Office, se não estiver usando um IDE como o Visual Studio, ou se precisar investigar um problema durante a execução do seu suplemento fora do IDE. Este artigo descreve como usar o Depurador a partir das ferramentas de desenvolvedor F12 no Windows 10 para testar seu suplemento do Office.

> [!NOTE]
> As instruções descritas neste artigo não podem ser usadas para depurar um suplemento do Outlook que usa funções Execute. Para depurar um suplemento do Outlook que usa funções Execute recomendamos que você use o Visual Studio no modo de script ou algum outro depurador de scripts.

## <a name="prerequisites"></a>Pré-requisitos

Você precisa dos seguintes softwares:

- As ferramentas do desenvolvedor F12, que estão incluídas no Windows 10. 
    
- O aplicativo cliente do Office que hospeda seu suplemento. 
    
- Seu suplemento. 

## <a name="using-the-debugger"></a>Utilização do depurador

Você pode usar o depurador das ferramentas de desenvolvedor F12 no Windows 10 para testar suplementos da AppSource ou suplementos que você adicionou de outros locais. Você pode iniciar as ferramentas de desenvolvedor F12 depois  que o suplemento estiver em execução. As ferramentas F12 abrem uma janela separada e não usam o Visual Studio.

> [!NOTE]
> O Depurador faz parte das ferramentas de desenvolvedor F12 no Windows 10 e no Internet Explorer. As versões anteriores do Windows não incluem o Depurador. 

Este exemplo usa o Word e um suplemento gratuito do AppSource.

1. Abra o Word e escolha um documento em branco. 
    
2. Na guia **Inserir** , no grupo Suplementos, escolha **Store** e selecione o suplemento **QR4Office** . (Você pode carregar qualquer suplemento da Store ou o seu catálogo de suplementos.)
    
3. Inicie as ferramentas de desenvolvimento F12 que correspondem à sua versão do Office:
    
   - Para a versão de 32 bits do Office, use C:\Windows\System32\F12\IEChooser.exe
    
   - Para a versão de 64 bits do Office, use C:\Windows\SysWOW64\F12\IEChooser.exe
    
   Quando você inicia IEChooser, uma janela separada denominada "Escolher destino para depurar" exibe os possíveis aplicativos para depurar. Selecione o aplicativo do seu interesse. Se você estiver escrevendo seu próprio suplemento, selecione o site onde você deseja ter o suplemento implantado, que pode ser uma URL de localhost. 
    
   Por exemplo, selecione **home.html**. 
    
   ![Tela do IEChooser, apontando para o suplemento de bolhas](../images/choose-target-to-debug.png)

4. Na janela F12, selecione o arquivo que você deseja depurar.
    
   Para selecionar o arquivo na janela F12, escolha o ícone de pasta acima do painel **script** (à esquerda). Na lista de arquivos disponíveis mostrados na lista suspensa, selecione **Home.js**.
    
5. Defina o ponto de interrupção.
    
   Para definir o ponto de interrupção em **home.js**, escolha a linha 144, na função `textChanged`. Você verá um ponto vermelho à esquerda da linha e uma linha correspondente no painel **Pilha de Chamadas e Pontos de Interrupção** (canto inferior direito). Para ver outras maneiras de definir um ponto de interrupção, confira [Inspecionar executando JavaScript com o Depurador](https://docs.microsoft.com/previous-versions/windows/internet-explorer/ie-developer/samples/dn255007(v=vs.85)). 
    
   ![Depurador com ponto de interrupção no arquivo home.js](../images/debugger-home-js-02.png)

6. Execute o suplemento para disparar o ponto de interrupção.
    
   No Word, escolha a caixa de texto URL na parte superior do painel de **QR4Office** e tente digitar um texto. No Depurador, no painel **Pilha de Chamadas e Pontos de Interrupção**, você verá que o ponto de interrupção foi disparado e mostra várias informações. Talvez você precise atualizar o Depurador para ver os resultados.
    
   ![Depurador com resultados do ponto de interrupção disparado](../images/debugger-home-js-01.png)


## <a name="see-also"></a>Confira também

- [Inspecionar executando JavaScript com o Depurador](https://docs.microsoft.com/previous-versions/windows/internet-explorer/ie-developer/samples/dn255007(v=vs.85))
- [Uso das ferramentas de desenvolvedor F12](https://docs.microsoft.com/previous-versions/windows/internet-explorer/ie-developer/samples/bg182326(v=vs.85))
