---
title: Publicar suplementos do Office usando a implantação centralizada por meio do centro de administração do Microsoft 365
description: Saiba como usar a implantação centralizada para implantar suplementos internos, bem como suplementos fornecidos por ISVs.
ms.date: 07/07/2020
localization_priority: Normal
ms.openlocfilehash: 0e99742be87b477b7c78295d08539de924f02466
ms.sourcegitcommit: 7ef14753dce598a5804dad8802df7aaafe046da7
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 07/10/2020
ms.locfileid: "45094250"
---
# <a name="publish-office-add-ins-using-centralized-deployment-via-the-microsoft-365-admin-center"></a>Publicar suplementos do Office usando a implantação centralizada por meio do centro de administração do Microsoft 365

O centro de administração do Microsoft 365 facilita para um administrador implantar os suplementos do Office para usuários e grupos dentro de sua organização. Os suplementos implantados por meio do Centro de administração ficam disponíveis imediatamente para os usuários nos aplicativos do Office, sem a necessidade de configuração do cliente. É possível usar a Implantação Centralizada para implantar suplementos internos, além de suplementos fornecidos por ISVs.

O centro de administração do Microsoft 365 atualmente oferece suporte aos seguintes cenários.

- Implantação Centralizada de suplementos novos e atualizados para usuários, grupos ou para uma organização.
- Implantação para várias plataformas de cliente, incluindo Windows, Mac e a Web. Para o Outlook, a implantação para iOS e Android também é suportada. (No entanto, enquanto a instalação do Excel, Outlook, Word e suplementos do PowerPoint no iPad é suportada, a implantação centralizada para iPad **não** é suportada.)
- Implantação no idioma inglês e para locatários no mundo inteiro.
- Implantação de suplementos hospedados na nuvem.
- Implantação de suplementos hospedados em um firewall.
- Implantação de suplementos do AppSource.
- Instalação automática de um suplemento para usuários que iniciam o aplicativo do Office.
- Remoção automática de um suplemento para os usuários se o administrador desativar ou excluir o suplemento ou se os usuários forem removidos do Azure Active Directory ou de um grupo no qual o suplemento foi implantado.

A implantação centralizada é a maneira recomendada para o administrador do Microsoft 365 implantar suplementos do Office em uma organização, desde que a organização atenda a todos os requisitos para usar a implantação centralizada. Para obter informações sobre como determinar se sua organização pode usar a implantação centralizada, consulte [determinar se a implantação centralizada de suplementos funciona para sua organização do Microsoft 365](/office365/admin/manage/centralized-deployment-of-add-ins).

> [!NOTE]
> Em um ambiente local sem conexão com o Microsoft 365, ou para implantar suplementos do SharePoint ou suplementos do Office que são direcionados para o Office 2013, use um [Catálogo de aplicativos do SharePoint](publish-task-pane-and-content-add-ins-to-an-add-in-catalog.md). Para implantar suplementos COM ou VSTO, use o Windows Installer ou o recurso ClickOnce, como descrito em [Implantando uma solução do Office](/visualstudio/vsto/deploying-an-office-solution).

## <a name="recommended-approach-for-deploying-office-add-ins"></a>Abordagem recomendada para implantar Suplementos do Office

Implante os suplementos do Office em fases para ajudar a garantir que a implantação corra bem. Recomendamos o plano a seguir:

1. Implante o suplemento em um pequeno conjunto de partes interessadas de negócios e membros do departamento de TI. Se a implantação for bem-sucedida, vá para a etapa 2.

2. Implante o suplemento para um conjunto maior de pessoas que usarão o suplemento dentro da empresa. Se a implantação for bem-sucedida, vá para a etapa 3.

3. Implante o suplemento para todo o conjunto de pessoas que usarão o suplemento.

Dependendo do tamanho do público-alvo, convém adicionar etapas a ou remover etapas deste procedimento.

## <a name="publish-an-office-add-in-via-centralized-deployment"></a>Publicar um suplemento por meio da Implantação Centralizada

Antes de começar, confirme se a sua organização atende a todos os requisitos para usar a implantação centralizada, conforme descrito em [determinar se a implantação centralizada de suplementos funciona para sua organização do Microsoft 365](/microsoft-365/admin/manage/centralized-deployment-of-add-ins).

Se sua organização atender aos requisitos, conclua as etapas a seguir para publicar um suplemento do Office por meio da Implantação Centralizada:

1. Entre no Microsoft 365 com sua conta corporativa ou de educação.
2. Selecione o ícone do inicializador de aplicativos no canto superior esquerdo e escolha **Administrador**.
3. No menu de navegação, pressione **Mostrar mais** e, em seguida, escolha **Configurações** > **Serviços e suplementos**.
4. Se você vir uma mensagem na parte superior da página anunciando o novo centro de administração do Microsoft 365, escolha a mensagem para ir para a visualização do centro de administração (consulte [About The Microsoft 365 Admin Center](/microsoft-365/admin/admin-overview/about-the-admin-center)).
5. Escolha **Implantar Suplemento** na parte superior da página.
6. Escolha **Avançar** depois de analisar os requisitos.
7. Escolha uma das opções a seguir na página **Implantação Centralizada**:

    - **Desejo adicionar um Suplemento da Office Store.**
    - **Tenho o arquivo de manifesto (.xml) neste dispositivo.** Para esta opção, escolha **Navegar** para localizar o arquivo de manifesto (.xml) que você deseja usar.
    - **Tenho uma URL para o arquivo de manifesto.** Para esta opção, digite a URL do manifesto no campo fornecido.

    ![Nova caixa de diálogo de suplemento no centro de administração do Microsoft 365](../images/new-add-in.png)

8. Se tiver selecionado a opção para adicionar um suplemento da Office Store, escolha o suplemento. É possível exibir suplementos disponíveis por meio das categorias **Sugeridos para você**, **Classificação** ou **Nome**. Você pode adicionar apenas suplementos gratuitos da Office Store. Atualmente não é possível adicionar suplementos pagos.

    > [!NOTE]
    > Com a opção da Office Store, as atualizações e os aprimoramentos do suplemento estão disponíveis automaticamente para usuários sem necessidade de intervenção.

    ![Selecionar uma caixa de diálogo de suplemento no centro de administração do Microsoft 365](../images/select-an-add-in.png)

9. Escolha **continuar** após analisar os detalhes do suplemento, política de privacidade e termos de licença.

    ![Página de suplemento selecionada no centro de administração do Microsoft 365](../images/selected-add-in-admin-center.png)

10. Na página **atribuir usuários** , escolha **todos**, **usuários/grupos específicos**ou **apenas eu**. Use a caixa Pesquisar para encontrar usuários e grupos para quem você quer implantar o suplemento. Para suplementos do Outlook, você também pode escolher o método de implantação **fixo**, **disponível**ou **opcional**.

    ![Gerenciar quem tem acesso e método de implantação no centro de administração do Microsoft 365](../images/manage-users-deployment-admin-center.png)

    > [!NOTE]
    > Um sistema de [logon único (SSO)](../develop/sso-in-office-add-ins.md) para suplementos está atualmente em versão prévia e não deve ser usado para suplementos de produção. Quando for implantado um suplemento que use o SSO, os usuários e grupos atribuídos também são compartilhados com suplementos que compartilham a mesma ID de aplicativo Azure. Todas as alterações nas atribuições do usuário também são aplicadas a esses suplementos. Os suplementos relacionados serão mostrados nessa página. Apenas em suplementos de SSO, essa página exibe a lista de permissões do Microsoft Graph exigida pelo suplemento.

11. Quando terminar, escolha **implantar**. Este processo pode levar até três minutos. Conclua a passo a passo, pressionando **Avançar**. Você verá o suplemento juntamente com outros aplicativos no Office 365.

    > [!NOTE]
    > Quando um administrador escolhe **implantar**, o consentimento é fornecido para todos os usuários.

    ![lista de aplicativos no centro de administração do Microsoft 365](../images/citations.png)

> [!TIP]
> Quando você implanta um novo suplemento para usuários e/ou grupos em sua organização, envie um email descrevendo quando e como usar o suplemento e incluindo links para conteúdo relevante da Ajuda, perguntas frequentes ou outros recursos de suporte.

## <a name="considerations-when-granting-access-to-an-add-in"></a>Considerações ao conceder acesso a um suplemento

Os administradores podem atribuir um suplemento a todos na organização ou a usuários e/ou grupos específicos de dentro da organização. A lista a seguir descreve as implicações de cada opção:

- **Todos**: como o nome sugere, essa opção atribui o suplemento a todos os usuários no locatário. Use essa opção com cautela e apenas para suplementos que sejam realmente universais para a sua organização.

- **Usuários**: Se você atribuir um suplemento a usuários individuais, será necessário atualizar as configurações da Central de implantação para o suplemento sempre que quiser atribuí-lo a outros usuários. Da mesma forma, será necessário atualizar as configurações da Central de implantação para o suplemento sempre que você quiser remover o acesso do usuário ao suplemento.

- **Grupos**: Se você atribuir um suplemento a um grupo, os usuários adicionados ao grupo serão atribuídos automaticamente ao suplemento. Da mesma forma, quando um usuário é removido de um grupo, ele automaticamente perde o acesso ao suplemento. Em ambos os casos, nenhuma ação adicional é necessária do Microsoft 365 admin.

Em geral, para facilitar a manutenção, recomendamos atribuir suplementos usando grupos sempre que possível. No entanto, em situações em que você deseja restringir o acesso do suplemento a um número muito pequeno de usuários, pode ser mais prático atribuir o suplemento a usuários específicos.

## <a name="add-in-states"></a>Estados de suplementos

A tabela a seguir descreve os estados diferentes de um suplemento.

|Estado|Como o estado ocorre|Impacto|
|-----|--------------------|------|
|**Ativo**|O administrador carregou o suplemento e o atribuiu a usuários e/ou grupos.|Os usuários e/ou grupos atribuídos ao suplemento o veem nos clientes do Office relevantes.|
|**Desativado**|O administrador desativou o suplemento.|Os usuários e/ou grupos atribuídos ao suplemento já não têm acesso a ele. Se o estado do suplemento for alterado de **Desativado** para **Ativo**, os usuários e os grupos recuperarão o acesso a ele.|
|**Excluído**|O administrador excluiu o suplemento.|Os usuários e/ou grupos atribuídos ao suplemento já não têm acesso a ele.|

## <a name="updating-office-add-ins-that-are-published-via-centralized-deployment"></a>Atualizar suplementos do Office que são publicados por meio de Implantação Centralizada

Depois de um suplemento do Office ter sido publicado por meio de Implantação Centralizada, as alterações feitas ao aplicativo Web do suplemento automaticamente estarão disponíveis para todos os usuários assim que as alterações forem implementadas no aplicativo Web. As alterações feitas a um [arquivo de manifesto XML](../develop/add-in-manifests.md) de um suplemento, por exemplo, para atualizar o ícone, texto ou comandos do suplemento, ocorrem da seguinte maneira:

- **Suplemento de linha de negócios**: se um administrador tiver carregado um arquivo de manifesto explicitamente ao implementar a implantação centralizada por meio do centro de administração do Microsoft 365, o administrador deverá carregar um novo arquivo de manifesto que contenha as alterações desejadas. Depois que o arquivo de manifesto atualizado for carregado, o suplemento será atualizado na próxima vez que os aplicativos relevantes do Office iniciarem.

  > [!NOTE]
  > Um administrador não precisa remover um suplemento de LOB para fazer uma atualização. Na seção suplementos, o administrador pode simplesmente escolher o suplemento de LOB e invocar essa funcionalidade pressionando o botão **Atualizar suplemento** presente no canto inferior direito.
  > 
  > ![A captura de tela mostra a caixa de diálogo atualizar suplemento no centro de administração do Microsoft 365](../images/update-add-in-admin-center.png)

- **Suplemento da Office Store**: se um administrador selecionou um suplemento da Office Store ao implementar a implantação centralizada por meio do centro de administração do Microsoft 365 e as atualizações de suplemento na Office Store, o suplemento será atualizado posteriormente via implantação centralizada. Na próxima vez que os aplicativos relevantes do Office iniciarem, o suplemento será atualizado.

## <a name="end-user-experience-with-add-ins"></a>Experiência do usuário final com suplementos

Depois que um suplemento tiver sido publicado por meio de Implantação Centralizada, os usuários finais podem começar a usá-lo em qualquer plataforma que o suplemento suporte.

Se o suplemento tiver suporte para comandos, eles serão exibidos na Faixa de Opções do Office a todos os usuários para os quais o suplemento for implantado. No exemplo a seguir, o comando **Pesquisar Citação** aparece na faixa de opções para o suplemento **Citações**.

![A captura de tela mostra uma seção da faixa de opções do aplicativo do Office com o comando de citação realçado no suplemento de citações](../images/search-citation.png)

Caso contrário, os usuários podem adicioná-lo ao aplicativo do Office da seguinte maneira:

1. No Word 2016 ou posterior, no Excel 2016 ou posterior ou no PowerPoint 2016 ou posterior, escolha **Inserir** > **Meus suplementos**.
2. Escolha **Administrador Gerenciado**, na janela do suplemento.
3. Escolha o suplemento e escolha **Adicionar**.

    ![A captura de tela mostra a guia Administração Gerenciada da página Suplementos do Office de um aplicativo do Office. O suplemento Citações é exibido na guia.](../images/office-add-ins-admin-managed.png)

No entanto, para o Outlook 2016 ou posterior, os usuários podem fazer o seguinte:

1. No Outlook, escolha **Página Inicial** > **Store**.
2. Escolha o item **Administrador Gerenciado** a guia do suplemento.
3. Escolha o suplemento e escolha **Adicionar**.

    ![A captura de tela mostra a área da página do Administrador Gerenciado da página da Store do aplicativo Outlook.](../images/outlook-add-ins-admin-managed.png)

## <a name="see-also"></a>Confira também

- [Determinar se a implantação centralizada de suplementos funciona para sua organização do Microsoft 365](/office365/admin/manage/centralized-deployment-of-add-ins)
