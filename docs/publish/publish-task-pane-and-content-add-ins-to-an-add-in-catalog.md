---
title: Publicar suplementos de painel de tarefas e de conteúdo em um catálogo de aplicativos do SharePoint
description: Para tornar os suplementos do Office acessíveis aos usuários em sua organização, os administradores podem carregar arquivos de manifesto dos suplementos do Office no catálogo de aplicativos da organização.
ms.date: 07/27/2021
ms.localizationpriority: medium
ms.openlocfilehash: 786fbd24790a1b8205fc3b0e8a15ce591cf66ca4
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 09/12/2021
ms.locfileid: "59152034"
---
# <a name="publish-task-pane-and-content-add-ins-to-a-sharepoint-app-catalog"></a>Publicar suplementos de painel de tarefas e de conteúdo em um catálogo de aplicativos do SharePoint

Um catálogo de aplicativos é um conjunto de sites dedicado em um aplicativo da Web do SharePoint ou uma locação do SharePoint Online que hospeda bibliotecas de documentos para suplementos do Office e do SharePoint. Para tornar os suplementos do Office acessíveis aos usuários em sua organização, os administradores podem carregar arquivos de manifesto dos Suplementos do Office no catálogo de aplicativos da organização. Quando um administrador registra um catálogo de aplicativos como um catálogo confiável, os usuários podem inserir o suplemento a partir da interface do usuário de inserção em um aplicativo cliente do Office.

> [!IMPORTANT]
>
> - Catálogos de aplicativo no SharePoint não oferecem suporte a recursos de suplemento que são implementados no nó `VersionOverrides` do [manifesto do suplemento](../develop/add-in-manifests.md), como por exemplo comandos de suplemento.
> - Se você estiver direcionando uma nuvem ou um ambiente híbrido, recomendamos que você [use](/microsoft-365/admin/manage/test-and-deploy-microsoft-365-apps) Aplicativos Integrados por meio do Centro de administração do Microsoft 365 para publicar seus complementos.
> - Catálogos de aplicativos no SharePoint não são compatíveis com o Office para Mac. Para implantar Suplementos do Office em clientes do Mac, envie-os para a [AppSource](/office/dev/store/submit-to-the-office-store).

## <a name="create-an-app-catalog"></a>Criar um catálogo de aplicativos

Conclua as etapas em uma das seções a seguir para criar um catálogo de aplicativos com o servidor SharePoint local ou no Microsoft 365.

### <a name="to-create-an-app-catalog-for-on-premises-sharepoint-server"></a>Para criar um catálogo de aplicativos para o SharePoint Server no local

Para criar o catálogo de aplicativos do SharePoint, siga as instruções em [Configurar o site do Catálogo de Aplicativos para um aplicativo da Web](/sharepoint/administration/manage-the-app-catalog).

Depois de criar o catálogo de aplicativos, siga as etapas para [publicar um Suplemento do Office](#publish-an-office-add-in).

### <a name="to-create-an-app-catalog-on-microsoft-365"></a>Para criar um catálogo de aplicativos Microsoft 365

Para criar o SharePoint de aplicativos, siga as instruções em [Create the App Catalog site collection](/sharepoint/use-app-catalog#step-1-create-the-app-catalog-site-collection). Depois de criar o catálogo de aplicativos, siga as etapas na próxima seção para publicar um Office Dedados.

## <a name="publish-an-office-add-in"></a>Publicar um Suplemento do Office

Conclua as etapas em uma das seções a seguir para publicar um Office Add-in em um catálogo de aplicativos no Microsoft 365 ou local SharePoint Server.

### <a name="to-publish-an-office-add-in-to-a-sharepoint-app-catalog-on-microsoft-365"></a>Para publicar um Office add-in em um catálogo SharePoint de aplicativos Microsoft 365

1. Vá para a [página Sites ativos do novo centro de administração do SharePoint](https://admin.microsoft.com/sharepoint?page=siteManagement&modern=true) e entre com uma conta que tenha [permissões de administrador](/sharepoint/sharepoint-admin-role) da sua organização.

    > [!NOTE]
    > Se você tiver Microsoft 365 [Alemanha,](https://go.microsoft.com/fwlink/p/?linkid=848041)entre no Centro de administração do Microsoft 365 , navegue até o centro de administração SharePoint e abra a página Mais recursos. <br>Se você tiver Microsoft 365 operado pela 21Vianet (China), entre no [Centro de administração do Microsoft 365](https://go.microsoft.com/fwlink/p/?linkid=850627), navegue até o centro de administração do SharePoint e abra a página Mais recursos.

1. Abra o site do catálogo de aplicativos selecionando sua URL na coluna URL.

    > [!NOTE]
    > Se você acabou de criar o site de catálogo de aplicativos na seção anterior, pode levar alguns minutos para o site concluir a configuração.

1. Escolha **Distribuir aplicativos para o Office**.
1. Na página **Aplicativos do Office**, escolha **Novo**.
1. Na caixa de diálogo **Adicionar um documento**, selecione o botão **Escolher Arquivos**.
1. Localize e especifique o arquivo [manifesto](../develop/add-in-manifests.md) para carregar e escolha **Abrir**.
1. Na caixa de diálogo **Adicionar um documento**, escolha **OK**.

### <a name="to-publish-an-add-in-to-an-app-catalog-with-on-premises-sharepoint-server"></a>Para publicar um suplemento em um catálogo de aplicativos com o SharePoint Server local

1. Abra a página da **Administração Central**.
1. No painel de tarefas à esquerda, escolha **Aplicativos**.
1. Na página **Aplicativos**, em **Gerenciamento de Aplicativos**, escolha **Gerenciar Catálogo de Aplicativos**.
1. Na página **Gerenciar Catálogo de Aplicativos**, verifique se você tem o aplicativo da Web correto selecionado no Seletor de **Aplicativos da Web**.
1. Escolha a URL sob a **URL do site** para abrir o site do catálogo de aplicativos.
1. Escolha **Distribuir aplicativos para o Office**.
1. Na página **Aplicativos do Office**, escolha **Novo**.
1. Na caixa de diálogo **Adicionar um documento**, selecione o botão **Escolher Arquivos**.
1. Localize e especifique o arquivo [manifesto](../develop/add-in-manifests.md) para carregar e escolha **Abrir**.
1. Na caixa de diálogo **Adicionar um documento**, escolha **OK**.

## <a name="insert-office-add-ins-from-the-app-catalog"></a>Inserir suplementos do Office do catálogo de aplicativos

Para aplicativos do Office online, você pode encontrar suplementos do Office no catálogo de aplicativos, concluindo as etapas a seguir.

1. Abra o aplicativo do Office online (Excel, PowerPoint ou Word).
1. Crie ou abra um documento.
1. Escolha **Inserir** > **Suplementos**.
1. Na caixa de diálogo Suplementos do Office, escolha a guia **MINHA ORGANIZAÇÃO**. Os Suplementos do Office estão listados.
1. Escolha um suplemento do Office e, em seguida, escolha **Adicionar**.

Para aplicativos do Office na área de trabalho, você pode encontrar suplementos do Office no catálogo de aplicativos concluindo as etapas a seguir.

1. Abra o aplicativo da área de trabalho do Office (Excel, Word ou PowerPoint)
1. Escolha **Arquivo** > **Opções** > **Central de Confiabilidade** > **Configurações da Central de Confiabilidade** > **Catálogos de Suplementos Confiáveis**.
1. Digite a URL do catálogo de aplicativos do SharePoint na caixa **URL do catálogo** e escolha **Adicionar catálogo**.
    Use a forma mais curta da URL. Por exemplo, se a URL do catálogo de aplicativos do SharePoint for:
    - `https://<domain>/sites/<AddinCatalogSiteCollection>/AgaveCatalog`

    Especifique somente a URL do conjunto de sites pai:
    - `https://<domain>/sites/<AddinCatalogSiteCollection>`
1. Feche e reabra o aplicativo do Office.
1. Escolha **Inserir** > **Obter Suplementos**.
1. Na caixa de diálogo Suplementos do Office, escolha a guia **MINHA ORGANIZAÇÃO**. Os Suplementos do Office estão listados.
1. Escolha um suplemento do Office e, em seguida, escolha **Adicionar**.

Como alternativa, um administrador pode especificar um catálogo de aplicativos no SharePoint usando a política de grupo. As configurações de política relevantes estão disponíveis nos arquivos de Modelo Administrativo [(ADMX/ADML) para Microsoft 365 Apps, Office 2019 e Office 2016](https://www.microsoft.com/download/details.aspx?id=49030) e são encontradas em Configuração do **Usuário\Políticas\Modelos Administrativos\Microsoft Office 2016\Segurança Configurações\Central de Confiação\Catálogos Confiáveis.**
