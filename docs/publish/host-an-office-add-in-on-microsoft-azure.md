---
title: Hospedar um suplemento do Office no Microsoft Azure | Microsoft Docs
description: Saiba como implantar o aplicativo Web de um suplemento no Azure e realizar sideload do suplemento para testar em um aplicativo cliente do Office.
ms.date: 07/07/2020
localization_priority: Normal
ms.openlocfilehash: a30f1a8219501a68e6f46f013ef46640a59fe4e9
ms.sourcegitcommit: 7ef14753dce598a5804dad8802df7aaafe046da7
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 07/10/2020
ms.locfileid: "45094229"
---
# <a name="host-an-office-add-in-on-microsoft-azure"></a>Hospedar um Suplemento do Office no Microsoft Azure

The simplest Office Add-in is made up of an XML manifest file and an HTML page. The XML manifest file describes the add-in's characteristics, such as its name, what Office desktop applications it can run in, and the URL for the add-in's HTML page. The HTML page is contained in a web app that users interact with when they install and run your add-in within an Office client application. You can host the web app of an Office Add-in on any web hosting platform, including Azure.

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
> In this walkthrough, you're using a local file share as a trusted catalog where you'll store the add-in XML manifest file. In a real-world scenario, you might instead choose to [deploy the XML manifest file to a SharePoint catalog](../publish/publish-task-pane-and-content-add-ins-to-an-add-in-catalog.md) or [publish the add-in to AppSource](/office/dev/store/submit-to-appsource-via-partner-center).

## <a name="step-2-add-the-file-share-to-the-trusted-add-ins-catalog"></a>Etapa 2:Adicionar o compartilhamento de arquivos ao catálogo de Suplementos Confiáveis

1. Inicie o Word e crie um documento.

    > [!NOTE]
    > Embora este exemplo use o Word, é possível usar qualquer aplicativo do Office que dê suporte a Suplementos do Office, como Excel, Outlook, PowerPoint ou Project.

2. Escolha **Arquivo** > **Opções**.

3. Na caixa de diálogo **Opções do Word**, escolha **Central de Confiabilidade**, depois **Configurações da Central de Confiabilidade**.

4. In the **Trust Center** dialog box, choose **Trusted Add-in Catalogs**. Enter the universal naming convention (UNC) path for the file share you created earlier as the **Catalog URL** (for example, \\\YourMachineName\AddinManifests), and then choose **Add catalog**. 

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
      
      - Choose the **Resource Group** for your site. If you create a new group, you also need to name it.
      
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

O Visual Studio cria um suplemento básico do Word que você pode publicar como está, sem fazer alterações no projeto da Web. Para criar um suplemento para outro tipo de host do Office, como o Excel, repita as etapas e escolha um tipo de projeto com o host do Office desejado.

## <a name="step-5-publish-your-office-add-in-web-app-to-azure"></a>Etapa 5: publicar seu aplicativo Web do suplemento do Office no Azure

1. Com seu projeto de suplemento aberto no Visual Studio, expanda o nó da solução no **Gerenciador de Soluções**, em seguida, selecione **Serviço de Aplicativo**.

2. Right-click the web project and then choose **Publish**. The web project contains Office Add-in web app files so this is the project that you publish to Azure.

3. Na guia **Publicar**:

      - Escolha **Serviço de Aplicativo do Microsoft Azure**.

      - Escolha **Selecionar Existentes**.

      - Escolha **Publicar**.

4. Visual Studio publishes the web project for your Office Add-in to your Azure web app. When Visual Studio finishes publishing the web project, your browser opens and shows a webpage with the text "Your App Service app has been created." This is the current default page for the web app.

5. Copie a URL raiz (por exemplo:https://YourDomain.azurewebsites.net); você precisará dela ao editar o arquivo de manifesto do suplemento, mais tarde neste artigo.

## <a name="step-6-edit-and-deploy-the-add-in-xml-manifest-file"></a>Etapa 6: Editar e implantar o arquivo de manifesto XML do suplemento

1. No Visual Studio, com o suplemento do Office de exemplo aberto no **Gerenciador de Soluções**, expanda a solução para que ambos os projetos sejam exibidos.

2. Expand the Office Add-in project (for example WordWebAddIn), right-click the manifest folder, and then choose **Open**. The add-in XML manifest file opens.

3. In the XML manifest file, find and replace all instances of "~remoteAppUrl" with the root URL of the add-in web app on Azure. This is the URL that you copied earlier after you published the add-in web app to Azure (for example: https://YourDomain.azurewebsites.net). 

4. Escolha **Arquivo** e **Salvar tudo**. Em seguida, copie o arquivo do manifesto XML (por exemplo, WordWebAddIn.xml).

5. Usando o programa **Gerenciador de Arquivos**, navegue até o compartilhamento de arquivos de rede que você criou na [Etapa 1: criar uma pasta compartilhada](../publish/host-an-office-add-in-on-microsoft-azure.md#step-1-create-a-shared-folder-to-host-your-add-in-xml-manifest-file) e cole o arquivo de manifesto na pasta.

## <a name="step-7-insert-and-run-the-add-in-in-the-office-client-application"></a>Etapa 7: Inserir e executar o suplemento no aplicativo cliente do Office

1. Inicie o Word e crie um documento.

2. Na faixa de opções, escolha **Inserir** > **Meus Suplementos**.

3. In the **Office Add-ins** dialog box, choose **SHARED FOLDER**. Word scans the folder that you listed as a trusted add-ins catalog (in [Step 2: Add the file share to the Trusted Add-ins catalog](../publish/host-an-office-add-in-on-microsoft-azure.md#step-2-add-the-file-share-to-the-trusted-add-ins-catalog)) and shows the add-ins in the dialog box. You should see an icon for your sample add-in.

4. Choose the icon for your add-in and then choose **Add**. A **Show Taskpane** button for your add-in is added to the ribbon.

5. On the ribbon of the **Home** tab, choose the **Show Taskpane** button. The add-in opens in a task pane to the right of the current document.

6. Verify that the add-in works by selecting some text in the document and choosing the **Highlight!** button in the task pane.

## <a name="see-also"></a>Confira também

- [Publicar seu Suplemento do Office](../publish/publish.md)
- [Publicar seu suplemento usando o Visual Studio](../publish/package-your-add-in-using-visual-studio.md)
