---
title: Realizar sideload de suplementos do Outlook para teste
description: Use o sideload para instalar um suplemento do Outlook para teste sem precisar primeiro colocá-lo em um catálogo de suplementos.
ms.date: 07/09/2020
localization_priority: Normal
ms.openlocfilehash: 9b44b988ddd6552d5f7d14088a0b6f3ae1e410ed
ms.sourcegitcommit: 7ef14753dce598a5804dad8802df7aaafe046da7
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 07/10/2020
ms.locfileid: "45093879"
---
# <a name="sideload-outlook-add-ins-for-testing"></a>Realizar sideload de suplementos do Outlook para teste

Você pode usar sideload para instalar um suplemento do Outlook para teste sem precisar primeiro colocá-lo em um catálogo de suplementos.

## <a name="sideload-an-add-in-in-outlook-on-the-web"></a>Realizar o sideload de um suplemento do Outlook na Web

O processo de Sideload de um suplemento no Outlook na Web depende se você está usando a versão nova ou clássica.

- Se sua barra de ferramentas de caixa de correio for parecida com a imagem a seguir, confira [Sideload de um suplemento no novo Outlook na Web](#sideload-an-add-in-in-the-new-outlook-on-the-web).

    ![captura de tela parcial da barra de ferramentas do novo Outlook na Web](../images/outlook-on-the-web-new-toolbar.png)

- Se sua barra de ferramentas de caixa de correio for parecida com a imagem a seguir, confira [Sideload de um suplemento no Outlook na Web clássico](#sideload-an-add-in-in-classic-outlook-on-the-web).

    ![captura de tela parcial da barra de ferramentas do Outlook na Web clássico](../images/outlook-on-the-web-classic-toolbar.png)

> [!NOTE]
> Se sua organização tiver incluído seu logotipo na barra de ferramentas da caixa de correio, você verá algo um pouco diferente do mostrado nas imagens anteriores.

### <a name="sideload-an-add-in-in-the-new-outlook-on-the-web"></a>Realizar sideload de um suplemento no novo Outlook na Web

1. Acesse o [Outlook no Office 365](https://outlook.office.com).

1. No Outlook na Web, crie uma nova mensagem.

1. Escolha **...** na parte inferior da nova mensagem e selecione **Obter Suplementos** menu que aparecer.

    ![Janela para redigir a mensagem no novo Outlook na Web com a opção Obter Suplementos realçada](../images/outlook-on-the-web-new-get-add-ins.png)

1. Na caixa de diálogo **Suplementos do Outlook**, selecione **Meus suplementos**.

    ![Suplementos para a caixa de diálogo do Outlook no novo Outlook na Web com Meus suplementos selecionado](../images/outlook-on-the-web-new-my-add-ins.png)

1. Localize a seção **Suplementos personalizados** no final da caixa de diálogo. Selecione o link **Adicionar um suplemento personalizado** e selecione **Adicionar do arquivo**.

    ![Captura de tela Gerenciar suplementos apontando para a opção Adicionar do arquivo](../images/outlook-sideload-desktop-add-from-file.png)

1. Locate the manifest file for your custom add-in and install it. Accept all prompts during the installation.

### <a name="sideload-an-add-in-in-classic-outlook-on-the-web"></a>Realizar sideload de um suplemento no Outlook na Web clássico

1. Acesse o [Outlook no Office 365](https://outlook.office.com).

1. Escolha o ícone de engrenagem na seção superior direita da barra de ferramentas e selecione **Gerenciar suplementos**.

    ![Captura de tela do Outlook na Web apontando para a opção Gerenciar suplementos](../images/outlook-sideload-web-manage-integrations.png)

1. Na página **Gerenciar suplementos**, selecione **Suplementos** e **Meus suplementos**.

    ![Caixa de diálogo da Loja do Outlook na Web com a opção Meus suplementos selecionada](../images/outlook-sideload-store-select-add-ins.png)

1. Localize a seção **Suplementos personalizados** no final da caixa de diálogo. Selecione o link **Adicionar um suplemento personalizado** e selecione **Adicionar do arquivo**.

    ![Captura de tela Gerenciar suplementos apontando para a opção Adicionar do arquivo](../images/outlook-sideload-desktop-add-from-file.png)

1. Locate the manifest file for your custom add-in and install it. Accept all prompts during the installation.

## <a name="sideload-an-add-in-in-outlook-on-the-desktop"></a>Realizar sideload de um suplemento do Outlook na área de trabalho

### <a name="outlook-2016-or-later"></a>Outlook 2016 ou posterior

1. Abra o Outlook 2016 ou posterior no Windows ou Mac.

1. Selecione o botão **Obter Suplementos** na faixa de opções.

    ![Faixa de opções do Outlook 2016 apontando para o botão Store](../images/outlook-sideload-desktop-store.png)

    > [!NOTE]
    > Caso não veja o botão **Obter Suplementos** em sua versão do Outlook, selecione o botão **Store** na faixa de opções.

1. Selecione **Suplementos** e, depois, **Meus suplementos**.

    ![Caixa de diálogo da Loja do Outlook 2016 com a opção Meus suplementos selecionada](../images/outlook-sideload-store-select-add-ins.png)

1. Localize a seção **Suplementos personalizados** no final da caixa de diálogo. Selecione o link **Adicionar um suplemento personalizado** e selecione **Adicionar do arquivo**.

    ![Captura de tela da Loja apontando para a opção Adicionar do arquivo](../images/outlook-sideload-desktop-add-from-file.png)

1. Locate the manifest file for your custom add-in and install it. Accept all prompts during the installation.

### <a name="outlook-2013"></a>Outlook 2013

1. Abra o Outlook 2013 no Windows.

1. Selecione o menu **arquivo** e, em seguida, selecione o botão **gerenciar suplementos** na guia **informações** . o Outlook abrirá um navegador.

1. Siga as etapas na seção [Sideload um suplemento no Outlook na Web](#sideload-an-add-in-in-outlook-on-the-web) de acordo com a sua versão do Outlook na Web.

## <a name="remove-a-sideloaded-add-in"></a>Remover um suplemento do suplementos foi feito

Para remover um suplemento do suplementos foi feito do Outlook, use as etapas descritas anteriormente neste artigo para localizar o suplemento na seção **suplementos personalizados** da caixa de diálogo que lista seus suplementos instalados. escolha as reticências ( `...` ) para o suplemento e, em seguida, escolha **remover** para remover o suplemento específico do.