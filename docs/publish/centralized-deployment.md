---
title: Publicar Office-ins usando a Implantação Centralizada por meio do Centro de administração do Microsoft 365
description: Saiba como usar a Implantação Centralizada para implantar os complementos internos, bem como os complementos fornecidos por ISVs.
ms.date: 03/22/2021
localization_priority: Normal
ms.openlocfilehash: b57e21f177fe66f03985ce6baee4d9eeda75d8bd
ms.sourcegitcommit: 883f71d395b19ccfc6874a0d5942a7016eb49e2c
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 07/09/2021
ms.locfileid: "53348787"
---
# <a name="publish-office-add-ins-using-centralized-deployment-via-the-microsoft-365-admin-center"></a>Publicar Office-ins usando a Implantação Centralizada por meio do Centro de administração do Microsoft 365

A Centro de administração do Microsoft 365 torna mais fácil para um administrador implantar Office Dedados para usuários e grupos em sua organização. Os suplementos implantados por meio do Centro de administração ficam disponíveis imediatamente para os usuários nos aplicativos do Office, sem a necessidade de configuração do cliente. Você pode usar a Implantação Centralizada para implantar suplementos internos, além de suplementos fornecidos por ISVs.

O Centro de administração do Microsoft 365 atualmente dá suporte aos seguintes cenários.

- Implantação Centralizada de suplementos novos e atualizados para usuários, grupos ou para uma organização.
- Implantação em várias plataformas de cliente, incluindo Windows, Mac e a Web. Para Outlook, a implantação para iOS e Android também é suportada. (No entanto, enquanto a instalação de Excel, Outlook, Word e PowerPoint do PowerPoint no iPad é suportada, a Implantação  Centralizada para iPad não é suportada.)
- Implantação no idioma inglês e para locatários no mundo inteiro.
- Implantação de suplementos hospedados na nuvem.
- Implantação de suplementos hospedados em um firewall.
- Implantação de suplementos do AppSource.
- Instalação automática de um suplemento para usuários que iniciam o aplicativo do Office.
- Remoção automática de um suplemento para os usuários se o administrador desativar ou excluir o suplemento ou se os usuários forem removidos do Azure Active Directory ou de um grupo no qual o suplemento foi implantado.

A Implantação Centralizada é a maneira recomendada para um administrador de Microsoft 365 implantar Office de uma organização, desde que a organização atenda a todos os requisitos para usar a Implantação Centralizada. Para obter informações sobre como determinar se sua organização pode usar a Implantação Centralizada, consulte [Determine if Centralized Deployment of add-ins works for your Microsoft 365 organization](/office365/admin/manage/centralized-deployment-of-add-ins).

> [!NOTE]
> Em um ambiente local sem conexão com o Microsoft 365, ou para implantar os SharePoint ou os complementos do Office destinados ao Office 2013, use um catálogo de aplicativos [SharePoint](publish-task-pane-and-content-add-ins-to-an-add-in-catalog.md). Para implantar suplementos COM ou VSTO, use o Windows Installer ou o recurso ClickOnce, como descrito em [Implantando uma solução do Office](/visualstudio/vsto/deploying-an-office-solution).

## <a name="recommended-approach-for-deploying-office-add-ins"></a>Abordagem recomendada para implantar Suplementos do Office

Implante os suplementos do Office em fases para ajudar a garantir que a implantação corra bem. Recomendamos o seguinte plano.

1. Implante o suplemento em um pequeno conjunto de partes interessadas de negócios e membros do departamento de TI. Se a implantação for bem-sucedida, vá para a etapa 2.

1. Implante o suplemento para um conjunto maior de pessoas que usarão o suplemento dentro da empresa. Se a implantação for bem-sucedida, vá para a etapa 3.

1. Implante o suplemento para todo o conjunto de pessoas que usarão o suplemento.

Dependendo do tamanho do público-alvo, convém adicionar etapas a ou remover etapas deste procedimento.

## <a name="publish-an-office-add-in-via-centralized-deployment"></a>Publicar um suplemento por meio da Implantação Centralizada

Antes de começar, confirme se sua organização atende a todos os requisitos para usar a Implantação Centralizada, conforme descrito em Determine if [Centralized Deployment of add-ins](/microsoft-365/admin/manage/centralized-deployment-of-add-ins)works for your Microsoft 365 organization .

Se a sua organização atender a todos os requisitos, conclua as etapas a seguir para publicar um Office de Office por meio da Implantação Centralizada.

1. Entre no Microsoft 365 com sua conta de trabalho ou educação.
1. Selecione o ícone do inicializador de aplicativos no canto superior esquerdo e escolha **Administrador**.
1. No menu de navegação, selecione **Mostrar mais** e, em seguida, escolha **Configurações**  >  **aplicativos integrados.**
1. Na parte superior da página, escolha **Complementos**.
1. Se você vir uma mensagem na parte superior da página anunciando o novo Centro de administração do Microsoft 365, escolha a mensagem para ir para a Visualização do Centro de Administração (consulte [Sobre](/microsoft-365/admin/admin-overview/about-the-admin-center)o Centro de administração do Microsoft 365 ).
1. Escolha **Implantar Suplemento** na parte superior da página.
1. Escolha **Avançar** depois de analisar os requisitos.
1. Escolha uma das seguintes opções na **página Implantação Centralizada.**

    - **Desejo adicionar um Suplemento da Office Store.**
    - **Tenho o arquivo de manifesto (.xml) neste dispositivo.** Para esta opção, escolha **Navegar** para localizar o arquivo de manifesto (.xml) que você deseja usar.
    - **Tenho uma URL para o arquivo de manifesto.** Para esta opção, digite a URL do manifesto no campo fornecido.

    ![Nova Add-In de diálogo no Centro de administração do Microsoft 365.](../images/new-add-in.png)

1. Se tiver selecionado a opção para adicionar um suplemento da Office Store, escolha o suplemento. É possível exibir suplementos disponíveis por meio das categorias **Sugeridos para você**, **Classificação** ou **Nome**. Você pode adicionar apenas suplementos gratuitos da Office Store. Atualmente não é possível adicionar suplementos pagos.

    > [!NOTE]
    > Com a opção da Office Store, as atualizações e os aprimoramentos do suplemento estão disponíveis automaticamente para usuários sem necessidade de intervenção.

    ![Selecione uma caixa de diálogo de complemento no Centro de administração do Microsoft 365.](../images/select-an-add-in.png)

1. Escolha **Continuar** após revisar os detalhes do complemento, Política de Privacidade e Termos de Licença.

    ![Página de complemento selecionada no Centro de administração do Microsoft 365.](../images/selected-add-in-admin-center.png)

1. Na página **Atribuir Usuários,** escolha **Todos**, **Usuários/Grupos Específicos** ou **Somente eu**. Use a caixa Pesquisar para encontrar usuários e grupos para quem você quer implantar o suplemento. Para Outlook, você também pode escolher o método de implantação **Fixed**, **Available** ou **Optional**.

    ![Gerenciar quem tem acesso e método de implantação Centro de administração do Microsoft 365.](../images/manage-users-deployment-admin-center.png)

    > [!NOTE]
    > Os complementos que utilizam o logom único [(SSO)](../develop/sso-in-office-add-ins.md) solicitarão que o administrador consenta com os escopos listados no manifesto do complemento.  Se o mesmo serviço de backing for usado em vários complementos (a mesma ID do Aplicativo do Azure é usada com o SSO em diferentes complementos), os escopos de cada complemento serão solicitados a consentir com cada implantação. Esta página também exibirá a lista de permissões que o complemento exige.

1. Quando terminar, escolha **Implantar**. Este processo pode levar até três minutos. Conclua a passo a passo, pressionando **Avançar**. Agora você vê o seu complemento juntamente com outros Office aplicativos.

    > [!NOTE]
    > Quando um administrador escolhe **Implantar**, o consentimento é dado para todos os usuários.

    ![Lista de aplicativos Centro de administração do Microsoft 365.](../images/citations.png)

> [!TIP]
> Quando você implanta um novo suplemento para usuários e/ou grupos em sua organização, envie um email descrevendo quando e como usar o suplemento e incluindo links para conteúdo relevante da Ajuda, perguntas frequentes ou outros recursos de suporte.

## <a name="considerations-when-granting-access-to-an-add-in"></a>Considerações ao conceder acesso a um suplemento

Os administradores podem atribuir um suplemento a todos na organização ou a usuários e/ou grupos específicos de dentro da organização. A lista a seguir descreve as implicações de cada opção.

- **Todos**: como o nome sugere, essa opção atribui o suplemento a todos os usuários no locatário. Use essa opção com cautela e apenas para suplementos que sejam realmente universais para a sua organização.

- **Usuários**: Se você atribuir um suplemento a usuários individuais, será necessário atualizar as configurações da Central de implantação para o suplemento sempre que quiser atribuí-lo a outros usuários. Da mesma forma, será necessário atualizar as configurações da Central de implantação para o suplemento sempre que você quiser remover o acesso do usuário ao suplemento.

- **Grupos**: Se você atribuir um suplemento a um grupo, os usuários adicionados ao grupo serão atribuídos automaticamente ao suplemento. Da mesma forma, quando um usuário é removido de um grupo, ele automaticamente perde o acesso ao suplemento. Em ambos os casos, nenhuma ação adicional é necessária do administrador Microsoft 365.

Em geral, para facilitar a manutenção, recomendamos atribuir suplementos usando grupos sempre que possível. No entanto, em situações em que você deseja restringir o acesso do suplemento a um número muito pequeno de usuários, pode ser mais prático atribuir o suplemento a usuários específicos.

## <a name="add-in-states"></a>Estados de suplementos

A tabela a seguir descreve os estados diferentes de um suplemento.

|Estado|Como o estado ocorre|Impacto|
|-----|--------------------|------|
|**Ativo**|O administrador carregou o suplemento e o atribuiu a usuários e/ou grupos.|Os usuários e/ou grupos atribuídos ao suplemento o veem nos clientes do Office relevantes.|
|**Desativado**|O administrador desativou o suplemento.|Os usuários e/ou grupos atribuídos ao suplemento já não têm acesso a ele. Se o estado do suplemento for alterado de **Desativado** para **Ativo**, os usuários e os grupos recuperarão o acesso a ele.|
|**Excluído**|O administrador excluiu o suplemento.|Os usuários e/ou grupos atribuídos ao suplemento já não têm acesso a ele.|

## <a name="updating-office-add-ins-that-are-published-via-centralized-deployment"></a>Atualizar suplementos do Office que são publicados por meio de Implantação Centralizada

Depois que um Office add-in tiver sido publicado por meio da Implantação Centralizada, todas as alterações feitas no aplicativo Web do complemento estarão disponíveis automaticamente para todos os usuários depois que essas alterações são implementadas no aplicativo Web. As alterações feitas no arquivo de manifesto [XML](../develop/add-in-manifests.md) de um complemento para, por exemplo, atualizar o ícone, o texto ou os comandos do complemento, ocorrem da seguinte forma:

- Complemento **de** linha de negócios : se um administrador carregou explicitamente um arquivo de manifesto (de seu dispositivo ou apontando para uma URL) ao implementar a Implantação Centralizada por meio do Centro de administração do Microsoft 365, o administrador deverá carregar um novo arquivo de manifesto que contenha as alterações desejadas. Depois que o arquivo de manifesto atualizado for carregado, o suplemento será atualizado na próxima vez que os aplicativos relevantes do Office iniciarem.

  > [!NOTE]
  > Um administrador não precisa remover um complemento LOB para fazer uma atualização. Na seção Add-ins, o administrador pode simplesmente escolher o complemento LOB e invocar essa funcionalidade pressionando o botão Atualizar o **add-in** presente no canto inferior direito.
  >
  > ![Captura de tela mostra a caixa de diálogo Atualizar o Centro de administração do Microsoft 365.](../images/update-add-in-admin-center.png)

- **Office Add-in** da Store: se um administrador selecionou um complemento da Loja do Office ao implementar a Implantação Centralizada por meio do Centro de administração do Microsoft 365 e as atualizações de complemento na loja do Office, o complemento será atualizado posteriormente por meio da Implantação Centralizada. Pode levar até 24 horas para que as atualizações de complemento da Loja fluam para todos os usuários finais. Após essa duração, na próxima vez que os aplicativos Office relevantes reiniciarem para esses usuários, o complemento será atualizado. Os usuários também podem disparar uma Atualização Manual para obter a versão mais recente do add-in da Loja selecionando **Inserir** Complementos de Tabulação Administrador  >    >  **Gerenciado**  >  **Guia Atualizar.**

## <a name="end-user-experience-with-add-ins"></a>Experiência do usuário final com suplementos

Depois que um suplemento tiver sido publicado por meio de Implantação Centralizada, os usuários finais podem começar a usá-lo em qualquer plataforma que o suplemento suporte.

Se o suplemento tiver suporte para comandos, eles serão exibidos na Faixa de Opções do Office a todos os usuários para os quais o suplemento for implantado. No exemplo a seguir, o comando **Pesquisar Citação** aparece na faixa de opções para o suplemento **Citações**.

![Captura de tela mostra uma seção da faixa Aplicativo do Office com o comando Citação de Pesquisa realçada no complemento Citações.](../images/search-citation.png)

Caso contrário, os usuários podem adicioná-lo ao aplicativo do Office da seguinte maneira:

1. No Word 2016 ou posterior, no Excel 2016 ou posterior ou no PowerPoint 2016 ou posterior, escolha **Inserir** > **Meus suplementos**.
1. Escolha **Administrador Gerenciado**, na janela do suplemento.
1. Escolha o suplemento e escolha **Adicionar**.

    ![A captura de tela mostra a guia Administração Gerenciada da página Suplementos do Office de um aplicativo do Office. O suplemento Citações é exibido na guia.](../images/office-add-ins-admin-managed.png)

No entanto, para o Outlook 2016 ou posterior, os usuários podem fazer o seguinte:

1. No Outlook, escolha **Página Inicial** > **Store**.
1. Escolha o item **Administrador Gerenciado** a guia do suplemento.
1. Escolha o suplemento e escolha **Adicionar**.

    ![A captura de tela mostra a área da página do Administrador Gerenciado da página da Store do aplicativo Outlook.](../images/outlook-add-ins-admin-managed.png)

## <a name="see-also"></a>Confira também

- [Determinar se a Implantação Centralizada de complementos funciona para sua organização do Microsoft 365](/office365/admin/manage/centralized-deployment-of-add-ins)
