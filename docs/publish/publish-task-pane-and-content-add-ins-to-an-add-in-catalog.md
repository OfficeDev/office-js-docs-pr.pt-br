---
title: Publicar suplementos de conteúdo e de painel de tarefas em um catálogo do SharePoint
description: Para disponibilizar os Suplementos do Office para os usuários na organização, os administradores podem carregar arquivos de manifesto de Suplementos do Office no catálogo de suplementos para uso em nas organizações deles.
ms.date: 05/22/2019
localization_priority: Priority
ms.openlocfilehash: bffbf3e83a2e6d8d0c63252c27ba54826611f78b
ms.sourcegitcommit: adaee1329ae9bb69e49bde7f54a4c0444c9ba642
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/24/2019
ms.locfileid: "34432240"
---
# <a name="publish-task-pane-and-content-add-ins-to-a-sharepoint-catalog"></a>Publicar suplementos de conteúdo e de painel de tarefas em um catálogo do SharePoint

Um catálogo de suplementos é um conjunto de sites dedicado em um aplicativo Web do SharePoint ou em locatário do SharePoint Online que hospeda bibliotecas de documentos para Suplementos do SharePoint e do Office. Para disponibilizar suplementos do Office nas empresas, os administradores podem carregar arquivos de manifesto de Suplementos do Office no catálogo de suplementos para uso em suas organizações. Quando um administrador registra um catálogo de suplementos como um catálogo confiável, os usuários podem inserir o suplemento a partir da interface de usuário em um aplicativo cliente do Office.

> [!IMPORTANT]
> - Os catálogos de suplementos no SharePoint não são compatíveis com recursos de suplementos implementados no nó `VersionOverrides` do [manifesto do suplemento](../develop/add-in-manifests.md), como comandos de suplemento.
> - Se você está direcionando para um ambiente híbrido ou de nuvem, recomendamos [usar a Implantação Centralizada por meio do Centro de Administração do Office 365](../publish/centralized-deployment.md) para publicar os suplementos.
> - Catálogos do SharePoint não são compatíveis com o Office para Mac. Para implantar Suplementos do Office em clientes do Mac, envie-os para a [AppSource](/office/dev/store/submit-to-the-office-store).   

## <a name="create-an-add-in-catalog"></a>Criação de um catálogo de suplementos

Conclua as etapas em uma das seções a seguir para criar um catálogo de suplementos no SharePoint ou no Office 365.

### <a name="to-create-an-add-in-catalog-for-on-premises-sharepoint"></a>Criar um catálogo de suplementos no SharePoint local.

> [!NOTE]
> A IU no SharePoint local ainda se refere aos suplementos como **aplicativos**.

1. Acesse o **Site da Administração Central**.

2. No painel de tarefas à esquerda, escolha os  **Aplicativos**.

3. Na página**Aplicativos**, em **Gerenciamento de Aplicativos**, escolha  **Gerenciar Catálogo de Aplicativos**.

4. Na página**Gerenciar Catálogo de Aplicativos**, verifique se você tem o aplicativo web correto selecionado no **Seletor de Aplicativo Web**.

5. Escolha  **Exibir configurações do site**.

6. Na página **Configurações do Site**, escolha **Administradores de conjunto de sites** para especificar os administradores de conjunto de sites e escolha **OK**.

7. Para conceder permissões de site aos usuários, escolha **Permissões de Site** e **Conceder Permissões**.

8. Na caixa de diálogo **Compartilhar "Site do Catálogo de Aplicativos"**, especifique um ou mais usuários do site, defina as permissões apropriadas, defina outras opções se for o caso e escolha **Compartilhar**.

9. Para adicionar suplementos ao catálogo de Suplementos do Office, escolha **Aplicativos do Office**.

### <a name="to-create-an-app-catalog-on-office-365"></a>Criar um catálogo de aplicativos no Office 365

Mesmo que o SharePoint nomeie um catálogo de "aplicativo", é possível registrar os Suplementos do Office no catálogo.

1. Vá para o centro de administração do Microsoft 365. Para saber mais sobre como encontrar o centro de administração, confira [Centro de administração do Microsoft 365](https://docs.microsoft.com/office365/admin/admin-overview/about-the-admin-center).

2. Na página do centro de administração do Microsoft 365, expanda a lista dos **Centros de administração** e selecione**SharePoint**.

    > [!NOTE]
    > Use o centro de administração do SharePoint Clássico para criar o catálogo. Se você estiver no novo centro de administração do SharePoint, escolha **Centro de administração do SharePoint clássico** no painel esquerdo.

3. No painel de tarefas à esquerda, escolha  **aplicativos**.

4. Na página **aplicativos**, escolha **Catálogo de Aplicativos**.
    > [!NOTE]
    > Se um catálogo de aplicativos já foi criado e exibido nesta página, você poderá ignorar o restante dessas etapas e ir para a próxima seção deste artigo para publicar o suplemento no catálogo.

5. Na página **Site do Catálogo de Aplicativo**, escolha **OK** para aceitar a opção padrão e criar um novo site de catálogo de suplementos.

6. Na página **Criar Conjunto de Sites do Catálogo de Aplicativos**, especifique o título do seu site de Catálogo de Aplicativos.

7. Especifique o **Endereço do site da Web**.

8. Especifique um **administrador**.

9. Defina a **Cota de Recursos de Servidor** como 0 (zero). (A cota de recursos de servidor está relacionada à limitação das soluções de área restrita com mau desempenho, mas não instala soluções de área restrita no seu site de catálogo de aplicativos.)

10. Escolha **OK**.

O catálogo de aplicativos foi criado.

## <a name="publish-an-add-in-to-an-app-catalog"></a>Publicar um suplemento em um catálogo de aplicativos

Para publicar um suplemento em um catálogo de aplicativo existente, conclua as etapas a seguir.

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

    Suplementos de conteúdo e de painel de tarefas neste catálogo agora ficam disponíveis na caixa de diálogo **Suplementos do Office**. Para acessá-los, escolha **Meus Suplementos** na guia **Inserir** e, em seguida, escolha **MINHA ORGANIZAÇÃO**.

## <a name="end-user-experience-with-the-add-in-catalog"></a>Experiência do usuário final com o catálogo de suplementos

Os usuários finais podem acessar o catálogo de suplementos em um aplicativo do Office realizando as seguintes etapas:

1. Em um aplicativo do Office, vá para **Arquivo** > **Opções** > **Central de Confiabilidade** > **Configurações da Central de Confiabilidade** > **Catálogos de Suplementos Confiáveis**.

2. Especifique a URL do _conjunto de sites do SharePoint pai_ do catálogo de suplementos. 

    Por exemplo, se a URL do catálogo de Suplementos do Office é:

    - `https:// _domain_ /sites/ _AddinCatalogSiteCollection_ /AgaveCatalog`

    Especifique somente a URL do conjunto de sites pai:

    - `https:// _domain_ /sites/ _AddinCatalogSiteCollection_`

3. Feche e reabra o aplicativo do Office. O catálogo de suplementos estará disponível na caixa de diálogo **Suplementos do Office**.

Como alternativa, um administrador pode especificar um catálogo de Suplementos do Office no SharePoint usando as políticas de grupo. Para saber mais, veja a seção [Usar uma Política de Grupo para gerenciar como os usuários podem instalar e usar os Suplementos do Office](/previous-versions/office/office-2013-resource-kit/jj219429(v=office.15)#using-group-policy-to-manage-how-users-can-install-and-use-apps-for-office).
