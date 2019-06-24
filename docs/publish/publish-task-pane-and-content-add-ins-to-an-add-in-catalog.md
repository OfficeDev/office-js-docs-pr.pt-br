---
title: Publicar suplementos de painel de tarefas e de conteúdo em um catálogo de aplicativos do SharePoint
description: Para tornar os suplementos do Office acessíveis aos usuários em sua organização, os administradores podem carregar arquivos de manifesto dos suplementos do Office no catálogo de aplicativos da organização.
ms.date: 06/20/2019
localization_priority: Priority
ms.openlocfilehash: 20b97855ce50e3f70e602f511882761c6fd80655
ms.sourcegitcommit: 382e2735a1295da914f2bfc38883e518070cec61
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 06/21/2019
ms.locfileid: "35128556"
---
# <a name="publish-task-pane-and-content-add-ins-to-a-sharepoint-app-catalog"></a>Publicar suplementos de painel de tarefas e de conteúdo em um catálogo de aplicativos do SharePoint

Um catálogo de aplicativos é um conjunto de sites dedicado em um aplicativo da Web do SharePoint ou uma locação do SharePoint Online que hospeda bibliotecas de documentos para suplementos do Office e do SharePoint. Para tornar os suplementos do Office acessíveis aos usuários em sua organização, os administradores podem carregar arquivos de manifesto dos Suplementos do Office no catálogo de aplicativos da organização. Quando um administrador registra um catálogo de aplicativos como um catálogo confiável, os usuários podem inserir o suplemento a partir da interface do usuário de inserção em um aplicativo cliente do Office.

> [!IMPORTANT]
> - Catálogos de aplicativo no SharePoint não oferecem suporte a recursos de suplemento que são implementados no nó `VersionOverrides` do [manifesto do suplemento](../develop/add-in-manifests.md), como por exemplo comandos de suplemento.
> - Se você está direcionando para um ambiente híbrido ou de nuvem, recomendamos [usar a Implantação Centralizada por meio do Centro de Administração do Office 365](../publish/centralized-deployment.md) para publicar os suplementos.
> - Catálogos de aplicativos no SharePoint não são compatíveis com o Office para Mac. Para implantar Suplementos do Office em clientes do Mac, envie-os para a [AppSource](/office/dev/store/submit-to-the-office-store).

## <a name="create-an-app-catalog"></a>Criar um catálogo de aplicativos

Conclua as etapas em uma das seções a seguir para criar um catálogo de aplicativos com o SharePoint Server local ou no Office 365.

### <a name="to-create-an-app-catalog-for-on-premises-sharepoint-server"></a>Para criar um catálogo de aplicativos para o SharePoint Server no local

Para criar o catálogo de aplicativos do SharePoint, siga as instruções em [Configurar o site do Catálogo de Aplicativos para um aplicativo da Web](https://docs.microsoft.com/pt-BR/sharepoint/administration/manage-the-app-catalog).

Depois de criar o catálogo de aplicativos, siga as etapas para [publicar um Suplemento do Office](#publish-an-office-add-in).

### <a name="to-create-an-app-catalog-on-office-365"></a>Criar um catálogo de aplicativos no Office 365

1. Vá para o centro de administração do Microsoft 365. Para saber mais sobre como encontrar o centro de administração, confira [Centro de administração do Microsoft 365](https://docs.microsoft.com/office365/admin/admin-overview/about-the-admin-center).

2. Na página do centro de administração do Microsoft 365, expanda a lista dos **Centros de administração** e selecione**SharePoint**.

    > [!NOTE]
    > Use o centro de administração do SharePoint Clássico para criar o catálogo. Se você estiver no novo centro de administração do SharePoint, escolha **Centro de administração do SharePoint clássico** no painel esquerdo.

3. No painel de tarefas à esquerda, escolha  **aplicativos**.

4. Na página **aplicativos**, escolha **Catálogo de Aplicativos**.
    > [!NOTE]
    > Se um catálogo de aplicativos já foi criado e exibido nesta página, você poderá ignorar o restante dessas etapas e ir para a próxima seção deste artigo para publicar o suplemento no catálogo.

5. Na página do **Site do Catálogo de Aplicativos**, escolha **OK** para aceitar a opção padrão e criar um novo site de catálogo de aplicativos.

6. Na página **Criar Conjunto de Sites do Catálogo de Aplicativos**, especifique o título do seu site de Catálogo de Aplicativos.

7. Especifique o **Endereço do site da Web**.

8. Especifique um **administrador**.

9. Defina a **Cota de Recursos de Servidor** como 0 (zero). (A cota de recursos de servidor está relacionada à limitação das soluções de área restrita com mau desempenho, mas não instala soluções de área restrita no seu site de catálogo de aplicativos.)

10. Escolha **OK**.

## <a name="publish-an-office-add-in"></a>Publicar um Suplemento do Office

Conclua as etapas em uma das seções a seguir para publicar um Suplemento do Office em um catálogo de aplicativos no Office 365 ou no SharePoint Server local.

### <a name="to-publish-an-office-add-in-to-a-sharepoint-app-catalog-on-office-365"></a>Para publicar um suplemento do Office em um catálogo de aplicativos do SharePoint no Office 365

1. Vá para o centro de administração do Microsoft 365. Para saber mais sobre como encontrar o centro de administração, confira [Centro de administração do Microsoft 365](https://docs.microsoft.com/office365/admin/admin-overview/about-the-admin-center).
2. Na página do centro de administração do Microsoft 365, expanda a lista dos **Centros de administração** e selecione**SharePoint**.
    > [!NOTE]
    > Use o centro de administração do SharePoint Clássico para criar o catálogo. Se você estiver no novo centro de administração do SharePoint, escolha **Centro de administração do SharePoint clássico** no painel esquerdo.
3. No painel de tarefas à esquerda, escolha  **aplicativos**.
4. Na página **aplicativos**, escolha **Catálogo de Aplicativos**.
5. Escolha **Distribuir aplicativos para o Office**.
6. Na página **Aplicativos do Office**, escolha **Novo**.
7. Na caixa de diálogo **Adicionar um documento**, selecione o botão **Escolher Arquivos**.
8. Localize e especifique o arquivo [manifesto](../develop/add-in-manifests.md) para carregar e escolha **Abrir**.
9. Na caixa de diálogo **Adicionar um documento**, escolha **OK**.

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

Como alternativa, um administrador pode especificar um catálogo de aplicativos no SharePoint usando a política de grupo. Para saber mais, veja a seção [Usar uma Política de Grupo para gerenciar como os usuários podem instalar e usar os Suplementos do Office](/previous-versions/office/office-2013-resource-kit/jj219429(v=office.15)#using-group-policy-to-manage-how-users-can-install-and-use-apps-for-office).
