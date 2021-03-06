---
title: Realizar sideload de suplementos do Outlook para teste
description: Use o sideload para instalar um suplemento do Outlook para teste sem precisar primeiro colocá-lo em um catálogo de suplementos.
ms.date: 02/10/2021
localization_priority: Normal
ms.openlocfilehash: b783b815af84a7fd8b4abd52cdd8e0925bfb9ecf
ms.sourcegitcommit: fefc279b85e37463413b6b0e84c880d9ed5d7ac3
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 02/12/2021
ms.locfileid: "50234244"
---
# <a name="sideload-outlook-add-ins-for-testing"></a>Realizar sideload de suplementos do Outlook para teste

Você pode usar sideload para instalar um suplemento do Outlook para teste sem precisar primeiro colocá-lo em um catálogo de suplementos.

## <a name="sideload-automatically"></a>Sideload automaticamente

Se você criou seu complemento do Outlook usando o gerador [Yeoman](https://github.com/OfficeDev/generator-office)para Os Complementos do Office, o sideload será feito da melhor maneira por meio da linha de comando. Isso aproveitará nossas ferramentas e sideload em todos os seus dispositivos com suporte em um único comando.

1. Usando a linha de comando, navegue até o diretório raiz do seu projeto de complemento gerado pelo Yeoman. Execute o comando `npm start`.

2. O seu complemento do Outlook fará o sideload automático para o Outlook em seu computador desktop. Você verá uma caixa de diálogo aparecer, informando que há uma tentativa de sideload do complemento, listando o nome e o local do arquivo de manifesto. Selecione **OK**, que registrará o manifesto.

> [!IMPORTANT]
> Se o manifesto contiver um erro ou o caminho para o manifesto for inválido, você receberá uma mensagem de erro.

3. Se o manifesto não contiver erros e o caminho for válido, seu complemento agora será sideloaded e estará disponível na área de trabalho e no Outlook na Web. Ele também será instalado em todos os seus dispositivos compatíveis.

## <a name="sideload-manually"></a>Realizar sideload manualmente

Embora seja recomendável realizar o sideload automaticamente pela linha de comando, conforme abordado na seção anterior, você também pode fazer o sideload manual de um complemento do Outlook com base no cliente do Outlook.

### <a name="outlook-on-the-web"></a>Outlook na Web

O processo de sideload de um complemento no Outlook na Web depende se você está usando a versão nova ou clássica.

- Se sua barra de ferramentas de caixa de correio for parecida com a imagem a seguir, confira [Sideload de um suplemento no novo Outlook na Web](#new-outlook-on-the-web).

    ![captura de tela parcial da barra de ferramentas do novo Outlook na Web](../images/outlook-on-the-web-new-toolbar.png)

- Se sua barra de ferramentas de caixa de correio for parecida com a imagem a seguir, confira [Sideload de um suplemento no Outlook na Web clássico](#classic-outlook-on-the-web).

    ![captura de tela parcial da barra de ferramentas do Outlook na Web clássico](../images/outlook-on-the-web-classic-toolbar.png)

> [!NOTE]
> Se sua organização tiver incluído seu logotipo na barra de ferramentas da caixa de correio, você verá algo um pouco diferente do mostrado nas imagens anteriores.

### <a name="new-outlook-on-the-web"></a>Novo Outlook na Web

1. Acesse o [Outlook na Web](https://outlook.office.com).

1. Crie uma nova mensagem.

1. Escolha **...** na parte inferior da nova mensagem e selecione **Obter Suplementos** menu que aparecer.

    ![Janela para redigir a mensagem no novo Outlook na Web com a opção Obter Suplementos realçada](../images/outlook-on-the-web-new-get-add-ins.png)

1. Na caixa de diálogo **Suplementos do Outlook**, selecione **Meus suplementos**.

    ![Suplementos para a caixa de diálogo do Outlook no novo Outlook na Web com Meus suplementos selecionado](../images/outlook-on-the-web-new-my-add-ins.png)

1. Localize a seção **Suplementos personalizados** no final da caixa de diálogo. Selecione o link **Adicionar um suplemento personalizado** e selecione **Adicionar do arquivo**.

    ![Captura de tela Gerenciar suplementos apontando para a opção Adicionar do arquivo](../images/outlook-sideload-desktop-add-from-file.png)

1. Localize o arquivo de manifesto de seu suplemento personalizado e instale-o. Aceite todos os prompts durante a instalação.

### <a name="classic-outlook-on-the-web"></a>Outlook na Web clássico

1. Acesse o [Outlook na Web](https://outlook.office.com).

1. Escolha o ícone de engrenagem na seção superior direita da barra de ferramentas e selecione **Gerenciar suplementos**.

    ![Captura de tela do Outlook na Web apontando para a opção Gerenciar suplementos](../images/outlook-sideload-web-manage-integrations.png)

1. Na página **Gerenciar suplementos**, selecione **Suplementos** e **Meus suplementos**.

    ![Caixa de diálogo da Loja do Outlook na Web com a opção Meus suplementos selecionada](../images/outlook-sideload-store-select-add-ins.png)

1. Localize a seção **Suplementos personalizados** no final da caixa de diálogo. Selecione o link **Adicionar um suplemento personalizado** e selecione **Adicionar do arquivo**.

    ![Captura de tela Gerenciar suplementos apontando para a opção Adicionar do arquivo](../images/outlook-sideload-desktop-add-from-file.png)

1. Localize o arquivo de manifesto de seu suplemento personalizado e instale-o. Aceite todos os prompts durante a instalação.

### <a name="outlook-on-the-desktop"></a>Outlook na área de trabalho

#### <a name="outlook-2016-or-later"></a>Outlook 2016 ou posterior

1. Abra o Outlook 2016 ou posterior no Windows ou Mac.

1. Selecione o botão **Obter Suplementos** na faixa de opções.

    ![Faixa de opções do Outlook 2016 apontando para o botão Obter Complementos](../images/outlook-sideload-desktop-store.png)

    > [!IMPORTANT]
    > Se você não vir o botão **Obter Complementos** em sua versão do Outlook, selecione:
    >
    > - **Botão Armazenar** na faixa de opções, se disponível.
    >
    >   OU
    >
    > - **Menu** Arquivo e, em **seguida, selecione**  o botão Gerenciar Complementos na guia Informações para abrir a caixa de diálogo **Add-ins** no Outlook na Web.<br>You can see more about the web experience in the previous section [Sideload an add-in in Outlook on the web](#outlook-on-the-web).

1. Se houver guias próximas à parte superior da caixa de diálogo, verifique se a guia **Add-ins** está selecionada. Escolha **Meus complementos.**

    ![Caixa de diálogo da Loja do Outlook 2016 com a opção Meus suplementos selecionada](../images/outlook-sideload-store-select-add-ins.png)

1. Localize a seção **Suplementos personalizados** no final da caixa de diálogo. Selecione o link **Adicionar um suplemento personalizado** e selecione **Adicionar do arquivo**.

    ![Captura de tela da Loja apontando para a opção Adicionar do arquivo](../images/outlook-sideload-desktop-add-from-file.png)

1. Localize o arquivo de manifesto de seu suplemento personalizado e instale-o. Aceite todos os prompts durante a instalação.

#### <a name="outlook-2013"></a>Outlook 2013

1. Abra o Outlook 2013 no Windows.

1. Selecione **o** menu Arquivo e selecione **o botão Gerenciar Complementos** na **guia** Informações. O Outlook abrirá a versão da Web em um navegador.

1. Siga as etapas na [seção Sideload](#outlook-on-the-web) de um complemento no Outlook na Web de acordo com sua versão do Outlook na Web.

## <a name="remove-a-sideloaded-add-in"></a>Remover um complemento de sideload

Em todas as versões do Outlook, a chave para remover um complemento de sideload é a caixa de diálogo Meus **Complementos,** que lista seus complementos instalados. Escolha as reellipses ( `...` ) para o complemento e selecione **Remover**.

Para navegar até a caixa de diálogo Meus **Complementos** para seu cliente do Outlook, use as últimas etapas listadas para [sideload manual](#sideload-manually) nas seções anteriores deste artigo.

Para remover um complemento de sideload do Outlook, use as etapas descritas anteriormente neste artigo para encontrar o add-in na seção de **complementos personalizados** da caixa de diálogo que lista seus complementos instalados. Escolha as reellipses ( ) para o complemento e, em seguida, escolha Remover para `...` remover esse complemento específico. 

