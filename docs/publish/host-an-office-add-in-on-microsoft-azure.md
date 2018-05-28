---
title: Hospedar um Suplemento do Office no Microsoft Azure
description: ''
ms.date: 01/25/2018
ms.openlocfilehash: f0d6a5a10d2ce0620b42566be03e2d36f8a922f2
ms.sourcegitcommit: c72c35e8389c47a795afbac1b2bcf98c8e216d82
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/23/2018
---
# <a name="host-an-office-add-in-on-microsoft-azure"></a>Hospedar um Suplemento do Office no Microsoft Azure

Os Suplementos do Office mais simples cont?m um arquivo de manifesto XML e uma p?gina HTML. O arquivo de manifesto XML descreve caracter?sticas do suplemento, como seu nome, quais aplicativos clientes do Office podem ser executados e a URL da p?gina HTML do suplemento. A p?gina HTML est? contida em um aplicativo Web com o qual os usu?rios interagem quando instalam e executam seu suplemento dentro de um aplicativo cliente do Office. Voc? pode hospedar o aplicativo Web de um suplemento do Office em qualquer plataforma de hospedagem Web, incluindo o Azure.

Este artigo descreve como implantar o aplicativo Web de um suplemento no Azure e [realizar sideload do suplemento](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md) para teste em um aplicativo cliente do Office.

## <a name="prerequisites"></a>Pr?-requisitos 

1. Instale o [Visual Studio 2017](https://www.visualstudio.com/downloads) e opte por incluir a carga de trabalho de **desenvolvimento do Azure**.

    > [!NOTE]
    > Se voc? tiver instalado o Visual Studio 2017 anteriormente, [use o Instalador do Visual Studio](https://docs.microsoft.com/en-us/visualstudio/install/modify-visual-studio) para garantir que a carga de trabalho de **desenvolvimento do Azure** esteja instalada. 

2. Instale o Office 2016. 
    
    > [!NOTE]
    > Se voc? ainda n?o tem o Office 2016, [registre-se para fazer uma avalia??o gratuita de um m?s](http://office.microsoft.com/en-us/try/?WT%2Eintid1=ODC%5FENUS%5FFX101785584%5FXT104056786).

3.  Obtenha uma assinatura do Azure.
    
    > [!NOTE]
    > Se voc? ainda n?o tem uma assinatura do Azure, pode [obter uma como parte da sua assinatura do MSDN](http://www.windowsazure.com/en-us/pricing/member-offers/msdn-benefits/) ou [registrar-se gratuitamente para uma avalia??o gratuita](https://azure.microsoft.com/en-us/pricing/free-trial). 

## <a name="step-1-create-a-shared-folder-to-host-your-add-in-xml-manifest-file"></a>Etapa 1: criar uma pasta compartilhada para hospedar o arquivo de manifesto XML do suplemento

1. Abra o Explorador de Arquivos em seu computador de desenvolvimento.
    
2. Clique com o bot?o direito do mouse na unidade C:\ e escolha **Novo** > **Pasta**.
    
3. Nomeie a nova pasta AddinManifests.
    
4. Clique com o bot?o direito do mouse na pasta AddinManifests e escolha **Compartilhar com** > **Pessoas espec?ficas**.
    
5. Em **Compartilhamento de Arquivos**, selecione a seta suspensa e escolha **Todos** > **Adicionar** > **Compartilhar**.
    
> [!NOTE]
> Nesta explica??o passo a passo, voc? est? usando um compartilhamento de arquivos local como um cat?logo confi?vel onde armazenar? o arquivo de manifesto XML do suplemento. Em um cen?rio real, em vez disso, ? poss?vel optar por [implantar o arquivo de manifesto XML a um cat?logo do SharePoint](../publish/publish-task-pane-and-content-add-ins-to-an-add-in-catalog.md) ou [publicar o suplemento no AppSource](https://docs.microsoft.com/en-us/office/dev/store/submit-to-the-office-store).

## <a name="step-2-add-the-file-share-to-the-trusted-add-ins-catalog"></a>Etapa 2:adicionar o compartilhamento de arquivos ao cat?logo de suplementos confi?veis

1.  Inicie o Word 2016 e crie um documento.

    > [!NOTE]
    > Embora este exemplo use o Word 2016, ? poss?vel usar qualquer aplicativo do Office que d? suporte a Suplementos do Office, como Excel, Outlook, PowerPoint ou Project 2016.
    
2.  Escolha **Arquivo**  >  **Op??es**.
    
3.  Na caixa de di?logo **Op??es do Word**, escolha **Central de Confiabilidade**, depois **Configura??es da Central de Confiabilidade**. 
    
4.  Na caixa de di?logo **Central de Confiabilidade**, escolha **Cat?logos de Suplementos Confi?veis**. Digite o caminho UNC (conven??o universal de nomenclatura) para o compartilhamento de arquivos que voc? criou anteriormente como a **URL do Cat?logo**. Por exemplo, \\\NomedoseuComputador\AddinManifests. Em seguida, escolha **Adicionar cat?logo**. 
    
5. Marque a caixa de sele??o **Mostrar no Menu**. 

    > [!NOTE]
    > Ao armazenar um arquivo de manifesto XML de suplemento em um compartilhamento especificado como um cat?logo de suplementos da Web confi?vel, o suplemento aparece em **Pasta Compartilhada** na caixa de di?logo **Suplementos do Office** quando o usu?rio navega at? a guia **Inserir** na faixa de op??es e escolhe **Meus Suplementos**.

6. Feche o Word 2016.

## <a name="step-3-create-a-web-app-in-azure"></a>Etapa 3: criar um aplicativo Web no Azure

Crie um aplicativo Web vazio no Azure usando [Visual Studio 2017](../publish/host-an-office-add-in-on-microsoft-azure.md#using-visual-studio-2017) ou o [portal do Azure](../publish/host-an-office-add-in-on-microsoft-azure.md#using-the-azure-portal).

### <a name="using-visual-studio-2017"></a>Usar o Visual Studio 2017

Para criar o aplicativo Web usando o Visual Studio 2017, realize as etapas a seguir.

1. No Visual Studio, no menu **Exibir**, escolha **Gerenciador de Servidores**. Clique com o bot?o direito do mouse em **Azure** e escolha **Conectar-se ? assinatura do Microsoft Azure**. Siga as instru??es para se conectar ? sua assinatura do Azure.
    
2. No Visual Studio, no **Gerenciador de Servidores**, expanda **Azure**, clique com o bot?o direito do mouse em **Servi?o de Aplicativo** e escolha **Criar novo aplicativo Web**.
    
3. Na caixa de di?logo **Criar Servi?o de Aplicativo**, forne?a estas informa??es:
    
      - Insira um **Nome do Aplicativo Web** exclusivo para seu site. O Azure verifica se o nome do site ? exclusivo em todo o dom?nio azurewebsites.net.

      - Escolha a **Assinatura** a ser usada para criar esse site.

      - Escolha o **Grupo de Recursos** para seu site. Se voc? criar um novo grupo, tamb?m precisar? dar um nome a ele.
    
      - Escolha o **Plano do Servi?o de Aplicativo** a ser usado para criar esse site. Se voc? criar um novo plano, tamb?m precisar? dar um nome a ele.
       
      - Escolha **Criar**.

    O novo aplicativo Web aparece no **Gerenciador de Servidores** em **Azure** >> **Servi?o de Aplicativo** >> (o grupo de recursos escolhido).
    
4. Clique com o bot?o direito do mouse no novo aplicativo Web e escolha **Exibir no Navegador**. O navegador ser? aberto e exibir? uma p?gina da Web com a mensagem "Seu aplicativo de Servi?o de Aplicativo foi criado".
    
5. Na barra de endere?os do navegador, altere a URL do aplicativo Web para que ela use HTTPS e pressione **Enter** para confirmar se o protocolo HTTPS foi habilitado. 

    > [!IMPORTANT]
    > [!include[HTTPS guidance](../includes/https-guidance.md)] Os sites do Azure fornecem automaticamente um ponto de extremidade HTTPS.
    
### <a name="using-the-azure-portal"></a>Usar o portal do Azure

Para criar o aplicativo Web usando o portal do Azure, realize as etapas a seguir.

1. Fa?a logon no [portal do Azure](https://portal.azure.com/) usando suas credenciais do Azure.
    
2. Escolha **Novo** > **Web + Celular** > **Aplicativo Web**. 

3. Na caixa de di?logo **Criar Aplicativo Web**, forne?a estas informa??es:
    
      - Insira um **Nome de aplicativo** exclusivo para seu site. O Azure verifica se o nome do site ? exclusivo em todo o dom?nio azureweb apps.net.

      - Escolha a **Assinatura** a ser usada para criar esse site.

      - Escolha o **Grupo de Recursos** para seu site. Se voc? criar um novo grupo, tamb?m precisar? dar um nome a ele.

      - Escolha o **SO** para seu site.
    
      - Escolha o **Plano do Servi?o de Aplicativo** a ser usado para criar esse site. Se voc? criar um novo plano, tamb?m precisar? dar um nome a ele.
       
      - Escolha **Criar**.

4. Escolha **Notifica??es** (o ?cone de sino localizado na borda superior do portal do Azure) e, em seguida, escolha a notifica??o **Implanta??es bem-sucedidas** para abrir a p?gina **Vis?o geral** no portal do Azure.

    > [!NOTE]
    > A notifica??o ser? alterada de **Implanta??o em andamento** para **Implanta??es bem-sucedidas** quando a implanta??o do site for conclu?da.

5. Na se??o **Fundamentos** da p?gina **Vis?o geral** do site no portal do Azure, escolha a URL exibida em **URL**. O navegador ser? aberto e exibir? uma p?gina da Web com a mensagem "Seu aplicativo de Servi?o de Aplicativo foi criado". 
    
6. Na barra de endere?os do navegador, altere a URL do aplicativo Web para que ela use HTTPS e pressione **Enter** para confirmar se o protocolo HTTPS foi habilitado. 

    > [!IMPORTANT]
    > [!include[HTTPS guidance](../includes/https-guidance.md)] Os sites do Azure fornecem automaticamente um ponto de extremidade HTTPS.

## <a name="step-4-create-an-office-add-in-in-visual-studio"></a>Etapa 4: criar um suplemento do Office no Visual Studio

1. Inicie o Visual Studio como um administrador.
    
2. Escolha **Arquivo** > **Novo** > **Projeto**.
    
3. Em **Modelos**, expanda **Visual C#** (ou **Visual Basic**), expanda **Office/SharePoint** e escolha **Suplementos**.
    
4. Escolha **Suplemento da Web do Word** e escolha **OK** para aceitar as configura??es padr?o.
       
O Visual Studio cria um suplemento b?sico do Word que voc? pode publicar como est?, sem fazer altera??es no projeto da Web.

## <a name="step-5-publish-your-office-add-in-web-app-to-azure"></a>Etapa 5: publicar seu aplicativo Web do suplemento do Office no Azure

1. Com seu projeto de suplemento aberto no Visual Studio, expanda o n? da solu??o no **Gerenciador de Solu??es** a fim de ver ambos os projetos para a solu??o.
    
2. Clique com bot?o direito do mouse no projeto da Web e escolha **Publicar**. O projeto da Web cont?m arquivos do aplicativo Web do suplemento do Office, portanto, esse ? o projeto que voc? publica no Azure.
    
3. Na guia **Publicar**:

      - Escolha **Servi?o de Aplicativo do Microsoft Azure**.
      
      - Escolha **Selecionar Existentes**.

      - Escolha **Publicar**. 

6. Na caixa de di?logo **Servi?o de Aplicativo**, localize e escolha o aplicativo Web que voc? criou na [Etapa 3: criar um aplicativo Web no Azure](../publish/host-an-office-add-in-on-microsoft-azure.md#step-3-create-a-web-app-in-azure) e, em seguida, escolha **OK**. 

    O Visual Studio publica o projeto da Web de seu Suplemento do Office no seu aplicativo Web do Azure. Quando o Visual Studio terminar de publicar o projeto da Web, o navegador abrir? e mostrar? uma p?gina da Web com o texto "Seu aplicativo de Servi?o de Aplicativo foi criado." Esta ? a p?gina padr?o atual do aplicativo Web.

7. Para ver a p?gina da Web do seu suplemento, altere o URL para que ele use HTTPS e especifique o caminho da p?gina HTML do seu suplemento (por exemplo: https://YourDomain.azurewebsites.net/Home.html). Isso confirma que o aplicativo Web do seu suplemento est? hospedado no Azure. Copie o URL raiz (por exemplo: https://YourDomain.azurewebsites.net); voc? precisar? dele ao editar o arquivo de manifesto do suplemento mais adiante neste artigo.
    
## <a name="step-6-edit-and-deploy-the-add-in-xml-manifest-file"></a>Etapa 6: editar e implantar o arquivo de manifesto XML do suplemento

1. No Visual Studio, com o suplemento do Office de exemplo aberto no **Gerenciador de Solu??es**, expanda a solu??o para que ambos os projetos sejam exibidos.
    
2. Expanda o projeto do Suplemento do Office (por exemplo, WordWebAddIn), clique com o bot?o direito do mouse na pasta do manifesto e escolha **Abrir**. O arquivo de manifesto XML do suplemento ? aberto.
    
3. No arquivo de manifesto XML, localize e substitua todas as inst?ncias de "~ remoteAppUrl" pelo URL raiz do aplicativo Web do suplemento no Azure. Esse ? o URL que voc? copiou anteriormente depois de publicar o aplicativo Web do suplemento no Azure (por exemplo: https://YourDomain.azurewebsites.net). 
    
4. Escolha **Arquivo** e **Salvar tudo**. Feche o arquivo de manifesto XML do suplemento.
    
5. No **Gerenciador de Solu??es**, clique com o bot?o direito do mouse na pasta do manifesto e escolha **Abrir Pasta no Gerenciador de Arquivos**.
    
6. Copie o arquivo de manifesto XML do suplemento (por exemplo, WordWebAddIn.xml). 
    
7. Navegue at? o compartilhamento de arquivos de rede que voc? criou na [Etapa 1: criar uma pasta compartilhada](../publish/host-an-office-add-in-on-microsoft-azure.md#step-1-create-a-shared-folder-to-host-your-add-in-xml-manifest-file) e cole o arquivo de manifesto na pasta.

## <a name="step-7-insert-and-run-the-add-in-in-the-office-client-application"></a>Etapa 7: inserir e executar o suplemento no aplicativo cliente do Office

1. Inicie o Word 2016 e crie um documento.
    
2. Na faixa de op??es, escolha **Inserir** > **Meus Suplementos**. 
    
3. Na caixa de di?logo **Suplementos do Office**, escolha **PASTA COMPARTILHADA**. O Word examina a pasta listada como um cat?logo de suplementos confi?veis (na [Etapa 2: adicionar o compartilhamento de arquivos ao cat?logo de suplementos confi?veis](../publish/host-an-office-add-in-on-microsoft-azure.md#step-2-add-the-file-share-to-the-trusted-add-ins-catalog)) e mostre os suplementos na caixa de di?logo. Voc? deve ver um ?cone de seu suplemento de exemplo.
    
4. Escolha o ?cone para seu suplemento e escolha **Adicionar**. Um bot?o **Mostrar Painel de Tarefas** para seu suplemento ? adicionado ? faixa de op??es. 

5. Na faixa de op??es da guia **P?gina Inicial**, escolha o bot?o **Mostrar Painel de Tarefas**. O suplemento ? aberto em um painel de tarefas ? direita do documento atual.
    
6. Para verificar se o suplemento funciona, selecione algum texto no documento e escolha o bot?o **Real?ar!** no painel de tarefas. 

## <a name="see-also"></a>Veja tamb?m

- [Publicar seu Suplemento do Office](../publish/publish.md)
- [Empacotar seu suplemento usando o Visual Studio para preparar a publica??o](../publish/package-your-add-in-using-visual-studio.md)
    
