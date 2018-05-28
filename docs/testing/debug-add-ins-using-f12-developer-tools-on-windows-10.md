---
title: Depurar suplementos usando as ferramentas de desenvolvedor F12 no Windows 10
description: ''
ms.date: 01/23/2018
ms.openlocfilehash: e1e4cde4a1a0fe27058346b93e8aaa39dd75a4e3
ms.sourcegitcommit: c72c35e8389c47a795afbac1b2bcf98c8e216d82
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/23/2018
---
# <a name="debug-add-ins-using-f12-developer-tools-on-windows-10"></a>Depurar suplementos usando as ferramentas de desenvolvedor F12 no Windows 10

As ferramentas de desenvolvedor F12 inclu?das no Windows 10 o ajudam a depurar, testar e acelerar suas p?ginas da Web. Voc? tamb?m pode us?-las para desenvolver e a depurar seu suplemento do Office se n?o estiver usando um IDE como o Visual Studio ou se precisar investigar um problema durante a execu??o do suplemento fora do IDE. Voc? pode iniciar as ferramentas de desenvolvedor F12 depois que o suplemento estiver em execu??o.

Este artigo mostra como ? poss?vel usar a ferramenta Depurador das ferramentas de desenvolvedor F12 no Windows 10 para testar seu suplemento do Office. ? poss?vel testar os suplementos do AppSource ou os suplementos que voc? adicionou de outros locais. As ferramentas F12 s?o exibidas em uma janela separada e n?o usam o Visual Studio.

> [!NOTE]
> O Depurador faz parte das ferramentas de desenvolvedor F12 no Windows 10 e no Internet Explorer. As vers?es anteriores do Windows n?o incluem o Depurador. 

## <a name="prerequisites"></a>Pr?-requisitos

Voc? precisa dos seguintes softwares:

- As ferramentas do desenvolvedor F12, que est?o inclu?das no Windows 10. 
    
- O aplicativo cliente do Office que hospeda seu suplemento. 
    
- Seu suplemento. 

## <a name="using-the-debugger"></a>Utiliza??o do depurador

Este exemplo usa o Word e um suplemento gratuito do AppSource.

1. Abra o Word e escolha um documento em branco. 
    
2. Na guia **Inserir**, no grupo Suplementos, escolha **Reposit?rio** e selecione o suplemento QR4Office. ? poss?vel carregar qualquer suplemento da Store ou seu cat?logo de suplementos.
    
3. Inicie as ferramentas de desenvolvimento F12 que correspondem ? sua vers?o do Office:
    
   - Para a vers?o de 32 bits do Office, use C:\Windows\System32\F12\F12Chooser.exe
    
   - Para a vers?o de 64 bits do Office, use C:\Windows\SysWOW64\F12\F12Chooser.exe
    
   Quando voc? inicia F12Chooser, uma janela separada denominada "Escolher destino para depurar" exibe os poss?veis aplicativos para depurar. Selecione o aplicativo do seu interesse. Se voc? estiver escrevendo seu pr?prio suplemento, selecione o site onde voc? deseja ter o suplemento implantado, que pode ser uma URL de localhost. 
    
   Por exemplo, selecione **home.html**. 
    
   ![Tela do F12Chooser, apontando para o suplemento bolhas](../images/choose-target-to-debug.png)

4. Na janela F12, selecione o arquivo que voc? deseja depurar.
    
   Para selecionar o arquivo, escolha o ?cone de pasta acima do painel **script** (? esquerda). A lista suspensa mostra os arquivos dispon?veis. Selecione home.js.
    
5. Defina o ponto de interrup??o.
    
   Para definir um ponto de interrup??o, escolha a linha 144, que est? na fun??o _textChanged_. Voc? ver? um ponto vermelho na parte esquerda da linha e uma linha correspondente no painel **Pilha de Chamadas e Pontos de Interrup??o** (parte inferior direita). Para ver outras formas de definir um ponto de interrup??o, confira [Inspecionar executando JavaScript com o Depurador](https://msdn.microsoft.com/library/dn255007%28v=vs.85%29.aspx). 
    
   ![Depurador com ponto de interrup??o no arquivo home.js](../images/debugger-home-js-02.png)

6. Execute o suplemento para acionar o ponto de interrup??o.
    
   Escolha a caixa de texto da URL na parte superior do painel QR4Office para alterar o texto. No Depurador, no painel **Pilha de Chamadas e Pontos de Interrup??o**, voc? ver? que o ponto de interrup??o est? disparado e mostra v?rias informa??es. Talvez voc? precise atualizar a ferramenta F12 para ver os resultados.
    
   ![Depurador com resultados do ponto de interrup??o disparado](../images/debugger-home-js-01.png)


## <a name="see-also"></a>Veja tamb?m

- [Inspecionar executando JavaScript com o Depurador](https://msdn.microsoft.com/library/dn255007%28v=vs.85%29.aspx)
- [Usando as ferramentas de desenvolvedor F12](https://msdn.microsoft.com/en-us/library/bg182326%28v=vs.85%29.aspx)
    
