---
title: Publicar suplementos de conteúdo e de painel de tarefas em um catálogo do SharePoint
description: Para disponibilizar os Suplementos do Office para os usuários na organização, os administradores podem carregar arquivos de manifesto de Suplementos do Office no catálogo de suplementos para uso em nas organizações deles.
ms.date: 03/19/2019
localization_priority: Priority
ms.openlocfilehash: ed4f9778e4cd7dccba00d2e8c019bd4441b70eeb
ms.sourcegitcommit: a2950492a2337de3180b713f5693fe82dbdd6a17
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/27/2019
ms.locfileid: "30870958"
---
# <a name="publish-task-pane-and-content-add-ins-to-a-sharepoint-catalog"></a>Publicar suplementos de conteúdo e de painel de tarefas em um catálogo do SharePoint

Um catálogo de suplementos é um conjunto de sites dedicado em um aplicativo Web do SharePoint ou em locatário do SharePoint Online que hospeda bibliotecas de documentos para Suplementos do SharePoint e do Office. Para disponibilizar suplementos do Office nas empresas, os administradores podem carregar arquivos de manifesto de Suplementos do Office no catálogo de suplementos para uso em suas organizações. Quando um administrador registra um catálogo de suplementos como um catálogo confiável, os usuários podem inserir o suplemento a partir da interface de usuário em um aplicativo cliente do Office.

> [!IMPORTANT]
> - Os catálogos de suplementos no SharePoint não são compatíveis com recursos de suplementos implementados no nó `VersionOverrides` do [manifesto do suplemento](../develop/add-in-manifests.md), como comandos de suplemento.
> - Se você está direcionando para um ambiente híbrido ou de nuvem, recomendamos [usar a Implantação Centralizada por meio do Centro de Administração do Office 365](../publish/centralized-deployment.md) para publicar os suplementos.
> - Catálogos do SharePoint não são compatíveis com o Office para Mac. Para implantar Suplementos do Office em clientes do Mac, envie-os para a [AppSource](/office/dev/store/submit-to-the-office-store).   

## <a name="set-up-an-add-in-catalog"></a>Configurar um catálogo de suplementos

Conclua as etapas em uma das seções a seguir para configurar um catálogo de suplementos no SharePoint ou no Office 365.

### <a name="to-set-up-an-add-in-catalog-for-on-premises-sharepoint"></a>Configurar um catálogo de suplementos no SharePoint local.

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

### <a name="to-set-up-an-add-in-catalog-on-office-365"></a>Configurar um catálogo de suplementos no Office 365

1. Na página do Centro de Administração do Office 365, escolha **Administrador** e **SharePoint**.

2. No painel de tarefas à esquerda, escolha **suplementos**.

3. Na página **suplementos**, escolha **Catálogo de Suplementos**.

4. Na página **Site do Catálogo de Suplementos**, escolha **OK** para aceitar a opção padrão e criar um novo site de catálogo de suplementos.

5. Na página **Criar Conjunto de Sites do Catálogo de Suplementos**, especifique o título do seu site de Catálogo de Suplementos.

6. Especifique o endereço do site da Web.

7. Defina a **cota de armazenamento** como o menor valor possível (atualmente 110). Você só instalará pacotes de suplementos neste conjunto de sites e eles são muito pequenos.

8. Defina a **Cota de Recursos de Servidor** como 0 (zero). (A cota de recursos de servidor está relacionada à limitação das soluções de área restrita com mau desempenho, mas não instala soluções de área restrita no seu site de catálogo de suplementos.)

9. Escolha **OK**.

10. Para adicionar um suplemento ao Site do Catálogo de Suplementos, navegue até o site que acabou de criar. No painel de navegação à esquerda, escolha **Suplementos do Office** e, para carregar um arquivo de manifesto do suplemento do Office, escolha **novo suplemento**.

## <a name="publish-an-add-in-to-an-add-in-catalog"></a>Publicar um suplemento em um catálogo de suplementos

Para publicar um suplemento em um catálogo suplementos, conclua as etapas a seguir.

1. Navegue até o catálogo de suplementos:

    - Abra a página principal Administração Central do SharePoint.

    - Selecione **Suplementos**.

    - Selecione **Gerenciar Catálogo de Suplementos**.

    - Escolha o link fornecido e escolha **Suplementos do Office** na barra de navegação à esquerda.

2. Escolha o link **Clique para adicionar um novo item**.

3. Escolha **Procurar** e especifique o [manifesto](../develop/add-in-manifests.md) para carregar.

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
