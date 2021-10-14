---
title: Hospedar um suplemento do Office no Microsoft Azure | Microsoft Docs
description: Saiba como implantar o aplicativo Web de um suplemento no Azure e realizar sideload do suplemento para testar em um aplicativo cliente do Office.
ms.date: 07/07/2020
ms.localizationpriority: medium
ms.openlocfilehash: af2079c4e6df38e64347ec0a67d441298758c41e
ms.sourcegitcommit: fb4a55764fb60e826ad06d15d1539e41df503b65
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 10/14/2021
ms.locfileid: "60356370"
---
# <a name="host-an-office-add-in-on-microsoft-azure"></a>Hospedar um Suplemento do Office no Microsoft Azure

Os Suplementos do Office mais simples contêm um arquivo de manifesto XML e uma página HTML. O arquivo de manifesto XML descreve as características do add-in, como seu nome Office, quais clientes de área de trabalho ele pode executar e a URL da página HTML do complemento. A página HTML está contida em um aplicativo Web com o qual os usuários interagem quando instalam e executam seu suplemento dentro de um aplicativo cliente do Office. Você pode hospedar o aplicativo Web de um suplemento do Office em qualquer plataforma de hospedagem Web, incluindo o Azure.

Este artigo descreve como implantar o aplicativo Web de um suplemento no Azure e [realizar sideload do suplemento](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md) para teste em um aplicativo cliente do Office.

## <a name="prerequisites"></a>Pré-requisitos 

1. Instale o [Visual Studio 2019](https://www.visualstudio.com/downloads) e opte por incluir a carga de trabalho de **desenvolvimento do Azure**.

    > [!NOTE]
    > Se você tiver instalado o Visual Studio 2019 anteriormente, [use o Instalador do Visual Studio](/visualstudio/install/modify-visual-studio) para garantir que a carga de trabalho de **desenvolvimento do Azure** esteja instalada. 

2. Instalar o Office.

    > [!NOTE]
    > Se você ainda não tem o Office, [registre-se para fazer uma avaliação gratuita de um mês](https://products.office.com/try?legRedir=true&WT.intid1=ODC_ENUS_FX101785584_XT104056786&CorrelationId=64c762de-7a97-4dd1-bb96-e231d7485735).

3. Obtenha uma assinatura do Azure.

    > [!NOTE]
    > Se você ainda não tem uma assinatura do Azure, pode [obter uma como parte da sua assinatura do Visual Studio](https://azure.microsoft.com/pricing/member-offers/visual-studio-subscriptions/) ou [registrar-se para uma avaliação gratuita](https://azure.microsoft.com/pricing/free-trial). 

## <a name="step-1-create-a-shared-folder-to-host-your-add-in-xml-manifest-file"></a>Etapa 1: criar uma pasta compartilhada para hospedar o arquivo de manifesto XML do suplemento

1. Abra o Explorador de Arquivos em seu computador de desenvolvimento.

2. Clique com o botão direito do mouse na unidade C:\ e escolha **Novo** > **Pasta**.

3. Nomeie a nova pasta AddinManifests.

4. Clique com o botão direito do mouse na pasta AddinManifests e escolha **Compartilhar com** > **Pessoas específicas**.

5. Em **Compartilhamento de Arquivos**, selecione a seta suspensa e escolha **Todos** > **Adicionar** > **Compartilhar**.

> [!NOTE]
> Nesta explicação passo a passo, você está usando um compartilhamento de arquivos local como um catálogo confiável onde armazenará o arquivo de manifesto XML do suplemento. Em um cenário real, em vez disso, é possível optar por [implantar o arquivo de manifesto XML a um catálogo do SharePoint](../publish/publish-task-pane-and-content-add-ins-to-an-add-in-catalog.md) ou [publicar o suplemento no AppSource](/office/dev/store/submit-to-appsource-via-partner-center).

## <a name="step-2-add-the-file-share-to-the-trusted-add-ins-catalog"></a>Etapa 2:Adicionar o compartilhamento de arquivos ao catálogo de Suplementos Confiáveis

1. Inicie o Word e crie um documento.

    > [!NOTE]
    > Embora este exemplo use o Word, é possível usar qualquer aplicativo do Office que dê suporte a Suplementos do Office, como Excel, Outlook, PowerPoint ou Project.

2. Escolha **Arquivo** > **Opções**.

3. Na caixa de diálogo **Opções do Word**, escolha **Central de Confiabilidade**, depois **Configurações da Central de Confiabilidade**.

4. Na caixa de diálogo **Central de Confiabilidade**, escolha **Catálogos de Suplementos Confiáveis**. Digite o caminho UNC (convenção universal de nomenclatura) para o compartilhamento de arquivos que você criou anteriormente como a **URL do Catálogo**. Por exemplo, \\\NomedoseuComputador\AddinManifests. Em seguida, escolha **Adicionar catálogo**. 

5. Marque a caixa de seleção **Mostrar no Menu**.

    > [!NOTE]
    > Ao armazenar um arquivo de manifesto XML de suplemento em um compartilhamento especificado como um catálogo de suplementos da Web confiável, o suplemento aparece em **Pasta Compartilhada** na caixa de diálogo **Suplementos do Office** quando o usuário navega até a guia **Inserir** na faixa de opções e escolhe **Meus Suplementos**.

6. Feche o Word.

## <a name="step-3-create-a-web-app-in-azure-using-the-azure-portal"></a>Etapa 3: Criar um aplicativo Web no Azure usando o portal do Azure

Para criar o aplicativo Web usando o portal do Azure, realize as etapas a seguir.

1. Faça logon no [portal do Azure](https://portal.azure.com/) usando suas credenciais do Azure.

2. Em **Serviços do Azure**, selecione **Aplicativos Web**.

3. Na página **Serviço de Aplicativo**, selecione **Adicionar**. Forneça estas informações:

      - Escolha a **Assinatura** a ser usada para criar esse site.
      
      - Escolha o **Grupo de Recursos** para seu site. Se você criar um novo grupo, também precisará dar um nome a ele.
      
      - Insira um **Nome de aplicativo** exclusivo para seu site. O Azure verifica se o nome do site é exclusivo em todo o domínio azureweb apps.net.

      - Escolha se deseja publicar usando um código ou um contêiner do docker.

      - Especificar uma **Pilha de tempo de execução**.

      - Escolha o **SO** para seu site.

      - Escolha uma **Região**.

      - Escolha o **Plano do Serviço de Aplicativo** a ser usado para criar esse site.

      - Escolha **Criar**.

4. A próxima página informa que a implantação está em andamento e quando ela é concluída. Quando estiver concluída, selecione **Ir ao recurso**.  

5. Na seção **Visão geral**, escolha a URL exibida em **URL**. O navegador será aberto e exibirá uma página da Web com a mensagem “Seu aplicativo de Serviço de Aplicativo está funcionando”.

    > [!IMPORTANT]
    > [!include[HTTPS guidance](../includes/https-guidance.md)] Os sites do Azure fornecem automaticamente um ponto de extremidade HTTPS.

## <a name="step-4-create-an-office-add-in-in-visual-studio"></a>Etapa 4: Criar um Suplemento do Office no Visual Studio

1. Inicie o Visual Studio como um administrador.

2. Escolha **Criar um novo projeto**.

3. Usando a caixa de pesquisa, insira **suplemento**.

4. Escolha **Suplemento da Web do Word** como o tipo de projeto e, em seguida, escolha **Avançar** para aceitar as configurações padrão.

O Visual Studio cria um suplemento básico do Word que você pode publicar como está, sem fazer alterações no projeto da Web. Para fazer um complemento para um aplicativo Office, como Excel, repita as etapas e escolha um tipo de projeto com seu aplicativo Office desejado.

## <a name="step-5-publish-your-office-add-in-web-app-to-azure"></a>Etapa 5: publicar seu aplicativo Web do suplemento do Office no Azure

1. Com seu projeto de suplemento aberto no Visual Studio, expanda o nó da solução no **Gerenciador de Soluções**, em seguida, selecione **Serviço de Aplicativo**.

2. Clique com botão direito do mouse no projeto da Web e escolha **Publicar**. O projeto da Web contém arquivos do aplicativo Web do suplemento do Office, portanto, esse é o projeto que você publica no Azure.

3. Na guia **Publicar**:

      - Escolha **Serviço de Aplicativo do Microsoft Azure**.

      - Escolha **Selecionar Existentes**.

      - Escolha **Publicar**.

4. O Visual Studio publica o projeto da Web de seu Suplemento do Office no seu aplicativo Web do Azure. Quando o Visual Studio terminar de publicar o projeto da Web, o navegador abrirá e mostrará uma página da Web com o texto "Seu aplicativo de Serviço de Aplicativo foi criado." Esta é a página padrão atual do aplicativo Web.

5. Copie a URL raiz (por exemplo: ); você precisará dela quando editar o arquivo de manifesto do complemento `https://YourDomain.azurewebsites.net` posteriormente neste artigo.

## <a name="step-6-edit-and-deploy-the-add-in-xml-manifest-file"></a>Etapa 6: Editar e implantar o arquivo de manifesto XML do suplemento

1. No Visual Studio, com o suplemento do Office de exemplo aberto no **Gerenciador de Soluções**, expanda a solução para que ambos os projetos sejam exibidos.

2. Expanda o projeto do Suplemento do Office (por exemplo, WordWebAddIn), clique com o botão direito do mouse na pasta do manifesto e escolha **Abrir**. O arquivo do manifesto XML do suplemento é aberto.

3. No arquivo do manifesto XML, localize e substitua todas as instâncias de "~remoteAppUrl" pela URL raiz do aplicativo web do suplemento no Azure. Esta é a URL que você copiou anteriormente depois de publicar o aplicativo Web do complemento para o Azure (por exemplo: `https://YourDomain.azurewebsites.net` ).

4. Escolha **Arquivo** e **Salvar tudo**. Em seguida, copie o arquivo do manifesto XML (por exemplo, WordWebAddIn.xml).

5. Usando o programa **Gerenciador de Arquivos**, navegue até o compartilhamento de arquivos de rede que você criou na [Etapa 1: criar uma pasta compartilhada](../publish/host-an-office-add-in-on-microsoft-azure.md#step-1-create-a-shared-folder-to-host-your-add-in-xml-manifest-file) e cole o arquivo de manifesto na pasta.

## <a name="step-7-insert-and-run-the-add-in-in-the-office-client-application"></a>Etapa 7: Inserir e executar o suplemento no aplicativo cliente do Office

1. Inicie o Word e crie um documento.

2. Na faixa de opções, escolha **Inserir** > **Meus Suplementos**.

3. Na caixa de diálogo **Suplementos do Office**, escolha **PASTA COMPARTILHADA**. O Word examina a pasta listada como um catálogo de suplementos confiáveis (na [Etapa 2: adicionar o compartilhamento de arquivos ao catálogo de suplementos confiáveis](../publish/host-an-office-add-in-on-microsoft-azure.md#step-2-add-the-file-share-to-the-trusted-add-ins-catalog)) e mostre os suplementos na caixa de diálogo. Você deve ver um ícone de seu suplemento de exemplo.

4. Escolha o ícone para seu suplemento e escolha **Adicionar**. Um botão **Mostrar Painel de Tarefas** para seu suplemento é adicionado à faixa de opções.

5. Na faixa de opções da guia **Página Inicial**, escolha o botão **Mostrar Painel de Tarefas**. O suplemento é aberto em um painel de tarefas à direita do documento atual.

6. Para verificar se o suplemento funciona, selecione algum texto no documento e escolha o botão **Realçar!** no painel de tarefas.

## <a name="see-also"></a>Confira também

- [Publicar seu Suplemento do Office](../publish/publish.md)
- [Publicar seu suplemento usando o Visual Studio](../publish/package-your-add-in-using-visual-studio.md)
