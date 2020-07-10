---
title: Publicar suplementos de painel de tarefas e de conteúdo em um catálogo de aplicativos do SharePoint
description: Para tornar os suplementos do Office acessíveis aos usuários em sua organização, os administradores podem carregar arquivos de manifesto dos suplementos do Office no catálogo de aplicativos da organização.
ms.date: 07/07/2020
localization_priority: Normal
ms.openlocfilehash: eabb60be927dc7fb274a0187a86f0c75592870bf
ms.sourcegitcommit: 7ef14753dce598a5804dad8802df7aaafe046da7
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 07/10/2020
ms.locfileid: "45094215"
---
# <a name="publish-task-pane-and-content-add-ins-to-a-sharepoint-app-catalog"></a>Publicar suplementos de painel de tarefas e de conteúdo em um catálogo de aplicativos do SharePoint

An app catalog is a dedicated site collection in a SharePoint web application or SharePoint Online tenancy that hosts document libraries for Office and SharePoint Add-ins. To make Office Add-ins accessible to users within their organization, administrators can upload Office Add-ins manifest files to the app catalog for their organization. When an administrator registers an app catalog as a trusted catalog, users can insert the add-in from the insertion UI in an Office client application.

> [!IMPORTANT]
> - Catálogos de aplicativo no SharePoint não oferecem suporte a recursos de suplemento que são implementados no nó `VersionOverrides` do [manifesto do suplemento](../develop/add-in-manifests.md), como por exemplo comandos de suplemento.
> - Se você estiver direcionando um ambiente híbrido ou de nuvem, recomendamos [usar a implantação centralizada por meio do centro de administração do Microsoft 365](../publish/centralized-deployment.md) para publicar seus suplementos.
> - Catálogos de aplicativos no SharePoint não são compatíveis com o Office para Mac. Para implantar Suplementos do Office em clientes do Mac, envie-os para a [AppSource](/office/dev/store/submit-to-the-office-store).

## <a name="create-an-app-catalog"></a>Criar um catálogo de aplicativos

Conclua as etapas em uma das seções a seguir para criar um catálogo de aplicativos com o SharePoint Server local ou no Office 365.

### <a name="to-create-an-app-catalog-for-on-premises-sharepoint-server"></a>Para criar um catálogo de aplicativos para o SharePoint Server no local

Para criar o catálogo de aplicativos do SharePoint, siga as instruções em [Configurar o site do Catálogo de Aplicativos para um aplicativo da Web](/sharepoint/administration/manage-the-app-catalog).

Depois de criar o catálogo de aplicativos, siga as etapas para [publicar um Suplemento do Office](#publish-an-office-add-in).

### <a name="to-create-an-app-catalog-on-microsoft-365"></a>Para criar um catálogo de aplicativos no Microsoft 365

Para criar o catálogo de aplicativos do SharePoint, siga as instruções em [criar o conjunto de sites do catálogo de aplicativos](/sharepoint/use-app-catalog#step-1-create-the-app-catalog-site-collection). Depois de criar o catálogo de aplicativos, siga as etapas na próxima seção para publicar um suplemento do Office.

## <a name="publish-an-office-add-in"></a>Publicar um Suplemento do Office

Conclua as etapas em uma das seções a seguir para publicar um suplemento do Office em um catálogo de aplicativos no Microsoft 365 ou no SharePoint Server local.

### <a name="to-publish-an-office-add-in-to-a-sharepoint-app-catalog-on-microsoft-365"></a>Para publicar um suplemento do Office em um catálogo de aplicativos do SharePoint no Microsoft 365

1. Vá para a [página Sites ativos do novo centro de administração do SharePoint](https://admin.microsoft.com/sharepoint?page=siteManagement&modern=true) e entre com uma conta que tenha [permissões de administrador](/sharepoint/sharepoint-admin-role) da sua organização.

>[!NOTE]
>Se você tiver o Microsoft 365 Alemanha, [entre no centro de administração do microsoft 365](https://go.microsoft.com/fwlink/p/?linkid=848041)e navegue até o centro de administração do SharePoint e abra a página mais recursos. <br>Se você tiver o Microsoft 365 operado pela 21Vianet (China), entre no centro [de administração do microsoft 365](https://go.microsoft.com/fwlink/p/?linkid=850627)e navegue até o centro de administração do SharePoint e abra a página mais recursos.
 
2. Abra o site do catálogo de aplicativos selecionando sua URL na coluna URL. 

>[!NOTE]
>Se você acabou de criar o site de catálogo de aplicativos na seção anterior, pode levar alguns minutos para que o site termine a configuração.

3. Escolha **Distribuir aplicativos para o Office**.
4. Na página **Aplicativos do Office**, escolha **Novo**.
5. Na caixa de diálogo **Adicionar um documento**, selecione o botão **Escolher Arquivos**.
6. Localize e especifique o arquivo [manifesto](../develop/add-in-manifests.md) para carregar e escolha **Abrir**.
7. Na caixa de diálogo **Adicionar um documento**, escolha **OK**.

### <a name="to-publish-an-add-in-to-an-app-catalog-with-on-premises-sharepoint-server"></a>Para publicar um suplemento em um catálogo de aplicativos com o SharePoint Server local

1. Abra a página da **Administração Central**.
2. No painel de tarefas à esquerda, escolha **Aplicativos**.
3. Na página **Aplicativos**, em **Gerenciamento de Aplicativos**, escolha **Gerenciar Catálogo de Aplicativos**.
4. Na página **Gerenciar Catálogo de Aplicativos**, verifique se você tem o aplicativo da Web correto selecionado no Seletor de **Aplicativos da Web**.
5. Escolha a URL sob a **URL do site** para abrir o site do catálogo de aplicativos.
6. Escolha **Distribuir aplicativos para o Office**.
7. Na página **Aplicativos do Office**, escolha **Novo**.
8. Na caixa de diálogo **Adicionar um documento**, selecione o botão **Escolher Arquivos**.
9. Localize e especifique o arquivo [manifesto](../develop/add-in-manifests.md) para carregar e escolha **Abrir**.
10. Na caixa de diálogo **Adicionar um documento**, escolha **OK**.

## <a name="insert-office-add-ins-from-the-app-catalog"></a>Inserir suplementos do Office do catálogo de aplicativos

Para aplicativos do Office online, você pode encontrar suplementos do Office no catálogo de aplicativos, concluindo as etapas a seguir.

1. Abra o aplicativo do Office online (Excel, PowerPoint ou Word).
2. Crie ou abra um documento.
3. Escolha **Inserir** > **Suplementos**.
4. Na caixa de diálogo Suplementos do Office, escolha a guia **MINHA ORGANIZAÇÃO**. Os Suplementos do Office estão listados.
5. Escolha um suplemento do Office e, em seguida, escolha **Adicionar**.

Para aplicativos do Office na área de trabalho, você pode encontrar suplementos do Office no catálogo de aplicativos concluindo as etapas a seguir.

1. Abra o aplicativo da área de trabalho do Office (Excel, Word ou PowerPoint)
2. Escolha **Arquivo** > **Opções** > ** Central de Confiabilidade** > **Configurações da Central de Confiabilidade** > **Catálogos de Suplementos Confiáveis**.
3. Digite a URL do catálogo de aplicativos do SharePoint na caixa **URL do catálogo** e escolha **Adicionar catálogo**.
    Use a forma mais curta da URL. Por exemplo, se a URL do catálogo de aplicativos do SharePoint for:
    - `https://<domain>/sites/<AddinCatalogSiteCollection>/AgaveCatalog`
    
    Especifique somente a URL do conjunto de sites pai:
    - `https://<domain>/sites/<AddinCatalogSiteCollection>`
4. Feche e reabra o aplicativo do Office.
5. Escolha **Inserir** > **Obter Suplementos**.
4. Na caixa de diálogo Suplementos do Office, escolha a guia **MINHA ORGANIZAÇÃO**. Os Suplementos do Office estão listados.
5. Escolha um suplemento do Office e, em seguida, escolha **Adicionar**.

Como alternativa, um administrador pode especificar um catálogo de aplicativos no SharePoint usando a política de grupo. As configurações de política relevantes estão disponíveis nos [arquivos de modelo administrativo (admx/adml) para os aplicativos do Microsoft 365, no office 2019 e no office 2016](https://www.microsoft.com/download/details.aspx?id=49030) e foram encontrados em **User. Administrativos\Microsoft Office 2016 \ segurança confiabilidade \ catálogos**.
