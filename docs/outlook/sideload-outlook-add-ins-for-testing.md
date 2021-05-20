---
title: Realizar sideload de suplementos do Outlook para teste
description: Use o sideload para instalar um suplemento do Outlook para teste sem precisar primeiro colocá-lo em um catálogo de suplementos.
ms.date: 05/13/2021
localization_priority: Normal
ms.openlocfilehash: 9d0fb246f6522c745658a09fce6934ee44d5079a
ms.sourcegitcommit: 693d364616b42eea66977eef47530adabc51a40f
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/19/2021
ms.locfileid: "52555189"
---
# <a name="sideload-outlook-add-ins-for-testing"></a>Realizar sideload de suplementos do Outlook para teste

Você pode usar sideload para instalar um suplemento do Outlook para teste sem precisar primeiro colocá-lo em um catálogo de suplementos.

## <a name="sideload-automatically"></a>Carga lateral automaticamente

Se você criou seu Outlook complemento usando [o gerador Yeoman para Office Add-ins,](https://github.com/OfficeDev/generator-office)o sideloading é melhor feito através da linha de comando. Isso aproveitará nossa ferramenta e carga lateral em todos os seus dispositivos suportados em um comando.

1. Usando a linha de comando, navegue até o diretório raiz do seu projeto de complemento gerado pela Yeoman. Execute o comando `npm start`.

1. O Outlook o complemento será automaticamente desviado para Outlook no computador de mesa. Você verá um diálogo aparecer, informando que há uma tentativa de carregar de lado o complemento, listando o nome e a localização do arquivo manifesto. Selecione **OK**, que registrará o manifesto.

    > [!IMPORTANT]
    > Se o manifesto contiver um erro ou o caminho para o manifesto for inválido, você receberá uma mensagem de erro.

1. Se o manifesto não contiver erros e o caminho for válido, o complemento agora será carregado lateralmente e disponível tanto na sua área de trabalho quanto em Outlook na web. Ele também será instalado em todos os seus dispositivos suportados.

## <a name="sideload-manually"></a>Carga lateral manualmente

Embora recomendemos fortemente o sideloading automaticamente através da linha de comando, conforme coberto na seção anterior, você também pode carregar manualmente um Outlook complemento com base no Outlook cliente.

### <a name="outlook-on-the-web"></a>Outlook na Web

O processo para carregar lateralmente um complemento Outlook na web depende se você está usando a versão nova ou clássica.

- Se sua barra de ferramentas de caixa de correio for parecida com a imagem a seguir, confira [Sideload de um suplemento no novo Outlook na Web](#new-outlook-on-the-web).

    ![captura de tela parcial da barra de ferramentas do novo Outlook na Web](../images/outlook-on-the-web-new-toolbar.png)

- Se sua barra de ferramentas de caixa de correio for parecida com a imagem a seguir, confira [Sideload de um suplemento no Outlook na Web clássico](#classic-outlook-on-the-web).

    ![captura de tela parcial da barra de ferramentas do Outlook na Web clássico](../images/outlook-on-the-web-classic-toolbar.png)

> [!NOTE]
> Se sua organização tiver incluído seu logotipo na barra de ferramentas da caixa de correio, você verá algo um pouco diferente do mostrado nas imagens anteriores.

### <a name="new-outlook-on-the-web"></a>Nova Outlook na web

1. Acesse o [Outlook na Web](https://outlook.office.com).

1. Crie uma nova mensagem.

1. Escolha **...** na parte inferior da nova mensagem e selecione **Obter Suplementos** menu que aparecer.

    ![Janela para redigir a mensagem no novo Outlook na Web com a opção Obter Suplementos realçada](../images/outlook-on-the-web-new-get-add-ins.png)

1. Na caixa de diálogo **Suplementos do Outlook**, selecione **Meus suplementos**.

    ![Suplementos para a caixa de diálogo do Outlook no novo Outlook na Web com Meus suplementos selecionado](../images/outlook-on-the-web-new-my-add-ins.png)

1. Localize a seção **Suplementos personalizados** no final da caixa de diálogo. Selecione o link **Adicionar um suplemento personalizado** e selecione **Adicionar do arquivo**.

    ![Captura de tela Gerenciar suplementos apontando para a opção Adicionar do arquivo](../images/outlook-sideload-desktop-add-from-file.png)

1. Localize o arquivo de manifesto de seu suplemento personalizado e instale-o. Aceite todos os prompts durante a instalação.

### <a name="classic-outlook-on-the-web"></a>Outlook clássica na web

1. Acesse o [Outlook na Web](https://outlook.office.com).

1. Escolha o ícone de engrenagem na seção superior direita da barra de ferramentas e selecione **Gerenciar suplementos**.

    ![Captura de tela do Outlook na Web apontando para a opção Gerenciar suplementos](../images/outlook-sideload-web-manage-integrations.png)

1. Na página **Gerenciar suplementos**, selecione **Suplementos** e **Meus suplementos**.

    ![Caixa de diálogo da Loja do Outlook na Web com a opção Meus suplementos selecionada](../images/outlook-sideload-store-select-add-ins.png)

1. Localize a seção **Suplementos personalizados** no final da caixa de diálogo. Selecione o link **Adicionar um suplemento personalizado** e selecione **Adicionar do arquivo**.

    ![Captura de tela Gerenciar suplementos apontando para a opção Adicionar do arquivo](../images/outlook-sideload-desktop-add-from-file.png)

1. Localize o arquivo de manifesto de seu suplemento personalizado e instale-o. Aceite todos os prompts durante a instalação.

### <a name="outlook-on-the-desktop"></a>Outlook no desktop

#### <a name="outlook-2016-or-later"></a>Outlook 2016 ou mais tarde

1. Abra Outlook 2016 ou posteriormente em Windows ou Mac.

1. Selecione o botão **Obter Suplementos** na faixa de opções.

    ![Outlook 2016 fita apontando para obter botão Add-ins](../images/outlook-sideload-desktop-store.png)

    > [!IMPORTANT]
    > Se você não ver o botão **Obter complementos na** sua versão de Outlook, selecione:
    >
    > - **Armazene** o botão na fita, se estiver disponível.
    >
    >   OU
    >
    > - **O** menu do arquivo e selecione o botão **Gerenciar complementos** na guia **Informações** para abrir a caixa de diálogo **Adicionar** em Outlook na web.<br>Você pode ver mais sobre a experiência da Web na seção anterior [Sideload um complemento Outlook na web](#outlook-on-the-web).

1. Se houver guias próximas à parte superior da caixa de diálogo, **certifique-se** de que a guia Adicionar está selecionada. Escolha **meus complementos**.

    ![Caixa de diálogo da Loja do Outlook 2016 com a opção Meus suplementos selecionada](../images/outlook-sideload-store-select-add-ins.png)

1. Localize a seção **Suplementos personalizados** no final da caixa de diálogo. Selecione o link **Adicionar um suplemento personalizado** e selecione **Adicionar do arquivo**.

    ![Captura de tela da Loja apontando para a opção Adicionar do arquivo](../images/outlook-sideload-desktop-add-from-file.png)

1. Localize o arquivo de manifesto de seu suplemento personalizado e instale-o. Aceite todos os prompts durante a instalação.

#### <a name="outlook-2013"></a>Outlook 2013

1. Aberto Outlook 2013 em Windows.

1. Selecione o menu **Arquivo** e selecione o botão **Gerenciar complementos** na guia **Informações.** Outlook abrirá a versão da Web em um navegador.

1. Siga os passos no [Sideload um complemento Outlook na](#outlook-on-the-web) seção web de acordo com sua versão de Outlook na web.

## <a name="remove-a-sideloaded-add-in"></a>Remova um complemento sideloaded

Em todas as versões de Outlook, a chave para remover um complemento com carga lateral é a caixa de diálogo **My Add-ins,** que lista seus complementos instalados. Escolha a elipse `...` () para o complemento e selecione **Remover**.

Para navegar até a caixa de diálogo **Meus Complementos** para o cliente Outlook, use as últimas etapas listadas para [sideloading manual](#sideload-manually) nas seções anteriores deste artigo.

Para remover um complemento sideloaded de Outlook, use as etapas descritas anteriormente neste artigo para encontrar o complemento na seção **de complementos personalizados** da caixa de diálogo que lista seus complementos instalados. Escolha a elipse `...` () para o complemento e escolha **Remover** para remover esse complemento específico. Feche a caixa de diálogo.
