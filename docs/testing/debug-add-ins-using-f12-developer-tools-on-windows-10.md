---
title: Depurar suplementos usando as ferramentas de desenvolvedor F12 no Windows 10
description: ''
ms.date: 01/23/2018
ms.openlocfilehash: 226773962fb1777a3a1f0e09445721ae2b8b5f5b
ms.sourcegitcommit: e1c92ba882e6eb03a165867c6021a6aa742aa310
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 08/20/2018
ms.locfileid: "22925602"
---
# <a name="debug-add-ins-using-f12-developer-tools-on-windows-10"></a>Depurar suplementos usando as ferramentas de desenvolvedor F12 no Windows 10

As ferramentas de desenvolvedor F12 incluídas no Windows 10 o ajudam a depurar, testar e acelerar suas páginas da Web. Você também pode usá-las para desenvolver e a depurar seu suplemento do Office se não estiver usando um IDE como o Visual Studio ou se precisar investigar um problema durante a execução do suplemento fora do IDE. Você pode iniciar as ferramentas de desenvolvedor F12 depois que o suplemento estiver em execução.

Este artigo mostra como é possível usar a ferramenta Depurador das ferramentas de desenvolvedor F12 no Windows 10 para testar seu suplemento do Office. É possível testar os suplementos do AppSource ou os suplementos que você adicionou de outros locais. As ferramentas F12 são exibidas em uma janela separada e não usam o Visual Studio.

> [!NOTE]
> O Depurador faz parte das ferramentas de desenvolvedor F12 no Windows 10 e no Internet Explorer. As versões anteriores do Windows não incluem o Depurador. 

## <a name="prerequisites"></a>Pré-requisitos

Você precisa dos seguintes softwares:

- As ferramentas do desenvolvedor F12, que estão incluídas no Windows 10. 
    
- O aplicativo cliente do Office que hospeda seu suplemento. 
    
- Seu suplemento. 

## <a name="using-the-debugger"></a>Utilização do depurador

Este exemplo usa o Word e um suplemento gratuito do AppSource.

1. Abra o Word e escolha um documento em branco. 
    
2. Na guia **Inserir**, no grupo Suplementos, escolha **Repositório** e selecione o suplemento QR4Office. É possível carregar qualquer suplemento da Store ou seu catálogo de suplementos.
    
3. Inicie as ferramentas de desenvolvimento F12 que correspondem à sua versão do Office:
    
   - Para a versão de 32 bits do Office, use C:\Windows\System32\F12\IEChooser.exe
    
   - Para a versão de 64 bits do Office, use C:\Windows\SysWOW64\F12\IEChooser.exe
    
   Quando você inicia o IEChooser, uma janela separada denominada "Escolher destino para depurar" exibe os possíveis aplicativos para depuração. Selecione o aplicativo do seu interesse. Se você estiver escrevendo seu próprio suplemento, selecione o site onde você deseja ter o suplemento implantado, que pode ser uma URL de localhost. 
    
   Por exemplo, selecione **home.html**. 
    
   ![Tela do IEChooser, apontando para o suplemento de bolhas](../images/choose-target-to-debug.png)

4. Na janela F12, selecione o arquivo que você deseja depurar.
    
   Para selecionar o arquivo, escolha o ícone de pasta acima do painel **script** (à esquerda). A lista suspensa mostra os arquivos disponíveis. Selecione home.js.
    
5. Defina o ponto de interrupção.
    
   Para definir um ponto de interrupção, escolha a linha 144, que está na função _textChanged_. Você verá um ponto vermelho na parte esquerda da linha e uma linha correspondente no painel **Pilha de Chamadas e Pontos de Interrupção** (parte inferior direita). Para ver outras formas de definir um ponto de interrupção, confira [Inspecionar executando JavaScript com o Depurador](https://docs.microsoft.com/previous-versions/windows/internet-explorer/ie-developer/samples/dn255007(v=vs.85)). 
    
   ![Depurador com ponto de interrupção no arquivo home.js](../images/debugger-home-js-02.png)

6. Execute o suplemento para acionar o ponto de interrupção.
    
   Escolha a caixa de texto da URL na parte superior do painel QR4Office para alterar o texto. No Depurador, no painel **Pilha de Chamadas e Pontos de Interrupção**, você verá que o ponto de interrupção está disparado e mostra várias informações. Talvez você precise atualizar a ferramenta F12 para ver os resultados.
    
   ![Depurador com resultados do ponto de interrupção disparado](../images/debugger-home-js-01.png)


## <a name="see-also"></a>Veja também

- [Inspecionar executando JavaScript com o Depurador](https://docs.microsoft.com/previous-versions/windows/internet-explorer/ie-developer/samples/dn255007(v=vs.85))
- [Usando as ferramentas de desenvolvedor F12](https://docs.microsoft.com/previous-versions/windows/internet-explorer/ie-developer/samples/bg182326(v=vs.85))
    
