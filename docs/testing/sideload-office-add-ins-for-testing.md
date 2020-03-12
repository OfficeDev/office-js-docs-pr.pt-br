---
title: Realizar sideload de suplementos do Office no Office na Web para teste
description: Testar o suplemento do Office no Office na web através de sideloading
ms.date: 02/18/2020
localization_priority: Normal
ms.openlocfilehash: 43dafda913516119ba1aadaae24701865596205f
ms.sourcegitcommit: 6c7c98f085dd20f827e0c388e672993412944851
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/06/2020
ms.locfileid: "42554682"
---
# <a name="sideload-office-add-ins-in-office-on-the-web-for-testing"></a>Realizar sideload de suplementos do Office no Office na Web para teste

Você pode instalar um suplemento do Office para teste usando sideloading, sem precisar primeiro colocá-lo em um catálogo de suplementos. O sideloading pode ser realizado no Office 365 ou no Office na Web. O procedimento é ligeiramente diferente nas duas plataformas.

Quando você realiza o sideload de um suplemento, o manifesto do suplemento é armazenado localmente do navegador e, portanto, se você limpar o cache do navegador ou alternar para um navegador diferente, precisará realizar o sideload do suplemento novamente.

> [!NOTE]
> A realização do sideload como descrito neste artigo tem suporte no Word, no Excel e no PowerPoint. Para realizar o sideload de um suplemento do Outlook, confira [Realizar sideload de suplementos do Outlook para teste](../outlook/sideload-outlook-add-ins-for-testing.md).

O vídeo a seguir oferece orientações para o processo de sideload do seu suplemento no Office na Web ou para área de trabalho.

> [!VIDEO https://www.youtube.com/embed/XXsAw2UUiQo]

## <a name="sideload-an-office-add-in-in-office-on-the-web"></a>Realizar sideload de um suplemento do Office no Office na Web

1. Abra o [Microsoft Office na Web](https://office.live.com/).

2. Em introdução **aos aplicativos online agora**, escolha **Excel**, **Word**ou **PowerPoint**; e, em seguida, abra um novo documento.

3. Abra a guia **Inserir** na faixa de opções e, na seção **suplementos** , escolha **suplementos do Office**.

4. Na caixa de diálogo **suplementos do Office** , selecione a guia **meus suplementos** , escolha **gerenciar meus suplementos**e, em seguida, **carregar meu suplemento**.

    ![A caixa de diálogo Suplementos do Office com um menu suspenso "Gerenciar meus suplementos" no canto superior direito e abaixo o menu suspenso com a opção "Carregar meu suplemento"](../images/office-add-ins-my-account.png)

5. **Navegue** até o arquivo de manifesto do suplemento e selecione **Carregar**.

    ![A caixa de diálogo Carregar suplemento com botões para pesquisar, carregar e cancelar.](../images/upload-add-in.png)

6. Verifique se o suplemento está instalado. Por exemplo, se for um comando do suplemento, ele deve aparecer na faixa de opções ou no menu de contexto. Se for um suplemento de painel de tarefas, o painel deve ser exibido.

> [!NOTE]
>Para testar o suplemento do Office com o Microsoft Edge, são necessárias duas etapas de configuração: 
>
> - Em um prompt de comando do Windows, execute a seguinte linha: `CheckNetIsolation LoopbackExempt -a -n="Microsoft.MicrosoftEdge_8wekyb3d8bbwe"`
>
> - Digite “**about:flags**” na barra de pesquisa do Microsoft Edge para exibir as opções de Configurações do Desenvolvedor.  Verifique a opção “**Permitir loopback do localhost**” e reinicie o Microsoft Edge.

>    ![A opção “Permitir loopback do localhost” do Microsoft Edge com a caixa marcada.](../images/allow-localhost-loopback.png)

## <a name="sideload-an-office-add-in-in-office-365"></a>Realizar sideload de um suplemento do Office no Office 365

1. Entre em sua conta do Office 365.

2. Abra o inicializador de aplicativos na extremidade esquerda da barra de ferramentas e selecione **Excel**, **Word**ou **PowerPoint**e, em seguida, crie um novo documento.

3. As etapas 3 a 6 são as mesmas da seção anterior **Realize sideload para um suplemento do Office no Office na Web**. 

## <a name="sideload-an-add-in-when-using-visual-studio"></a>Sideload de um suplemento usando o Visual Studio

Se estiver usando o Visual Studio para desenvolver o suplemento, o processo de sideload é semelhante. A única diferença é que você deve atualizar o valor do elemento **SourceURL** no manifesto para incluir a URL completa em que o suplemento for implantado.

> [!NOTE]
> Embora você possa realizar o sideload de suplementos do Visual Studio para o Office na Web, não é possível depurá-los no Visual Studio. Para depurar você precisará usar as ferramentas de depuração do navegador. Para saber mais, confira [Depurar suplementos no Office na Web](debug-add-ins-in-office-online.md).

1. No Visual Studio, abra a janela **Propriedades** escolhendo **Modo de exibição** -> **Janela de propriedades**.
2. No **Gerenciador de Soluções**, selecione o projeto Web. Isso exibirá as propriedades para o projeto na janela **Propriedades**.
3. Na janela Propriedades, copie a **URL de SSL**.
4. No projeto de suplemento, abra o arquivo XML do manifesto. Certifique-se de que você está editando o XML do código-fonte. Para alguns tipos de projeto o Visual Studio abrirá o modo de exibição de visualização do XML que não funcionará para a próxima etapa.
5. Pesquisar e substituir todas as instâncias de **~remoteAppUrl/** pela URL de SSL que você copiou. Você verá várias substituições dependendo do tipo de projeto e as novas URLs serão muito similares a `https://localhost:44300/Home.html`.
6. Salve o arquivo XML.
7. Clique com botão direito do mouse no projeto Web e escolha **Depurar** -> **Iniciar nova instância**. Isso executará o projeto Web sem iniciar o Office.
8. No Office na Web, realize o sideload do suplemento usando as etapas descritas anteriormente em [Sideload de um suplemento do Office no Office na Web](#sideload-an-office-add-in-in-office-on-the-web).

## <a name="remove-a-sideloaded-add-in"></a>Remover um suplemento do suplementos foi feito

Você pode remover um suplemento suplementos foi feito anteriormente limpando o cache do navegador. Além disso, se você fizer alterações no manifesto do suplemento (por exemplo, atualizar nomes de arquivo de ícones ou texto de comandos de suplemento), talvez seja necessário limpar o cache e, em seguida, resideload o suplemento usando o manifesto atualizado. Isso permitirá que o Office processe o suplemento conforme descrito no manifesto atualizado.
