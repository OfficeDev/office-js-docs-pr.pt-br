# <a name="publish-office-add-ins-using-centralized-deployment-via-the-office-365-admin-center"></a>Publicar Suplementos do Office usando a Implantação Centralizada por meio do Centro de administração do Office 365

No Centro de administração do Office 365, é mais fácil para o administrador implantar Suplementos do Office para usuários e grupos dentro da organização. Os suplementos implantados por meio do Centro de administração ficam disponíveis imediatamente para os usuários nos aplicativos do Office, sem a necessidade de configuração do cliente. Você pode usar a Implantação Centralizada para implantar suplementos internos, além de suplementos fornecidos por ISVs.

Atualmente, o Centro de administração do Office 365 tem suporte para os seguintes cenários:

- Implantação Centralizada de suplementos novos e atualizados para usuários, grupos ou para uma organização.
- Implantação para várias plataformas, inclusive Windows e Office Online; em breve para Mac.
- Implantação no idioma inglês e para locatários no mundo inteiro.
- Implantação de suplementos hospedados na nuvem.
- Implantação de suplementos que são hospedados em um firewall.
- Implantação de suplementos da Office Store.
- Instalação automática de um suplemento para usuários que iniciam o aplicativo do Office.
- Remoção automática de um suplemento para os usuários se o administrador desativar ou excluir o suplemento ou se os usuários forem removidos do Azure Active Directory ou de um grupo no qual o suplemento foi implantado.

A Implantação Centralizada é a maneira recomendada para o administrador do Office 365 implantar Suplementos do Office em uma organização, desde que a organização atenda a todos os requisitos para usar a Implantação Centralizada. Confira informações sobre como determinar se sua organização pode usar a Implantação Centralizada em [Determinar se a Implantação Centralizada de suplementos funciona para sua organização do Office 365](https://support.office.com/pt-BR/article/Determine-if-Centralized-Deployment-of-add-ins-works-for-your-Office-365-organization-B4527D49-4073-4B43-8274-31B7A3166F92).

>**Observação:** Em um ambiente local sem conexão com o Office 365, ou para implantar Suplementos do SharePoint ou Office que visam o Office 2013, use um [Catálogo de suplementos do SharePoint](publish-task-pane-and-content-add-ins-to-an-add-in-catalog.md). Para implantar suplementos COM ou VSTO, use o Windows Installer ou o recurso ClickOnce conforme descrito em [Implantando uma solução do Office](https://msdn.microsoft.com/pt-BR/library/bb386179.aspx).

## <a name="recommended-approach-for-deploying-office-add-ins"></a>Abordagem recomendada para implantar Suplementos do Office

Implante os suplementos do Office em fases para ajudar a garantir que a implantação corra bem. Recomendamos o plano a seguir:

1. Implante o suplemento em um pequeno conjunto de partes interessadas de negócios e membros do departamento de TI. Se a implantação for bem-sucedida, vá para a etapa 2.

2. Implante o suplemento para um conjunto maior de pessoas que usarão o suplemento dentro da empresa. Se a implantação for bem-sucedida, vá para a etapa 3.

3. Implante o suplemento para todo o conjunto de pessoas que usarão o suplemento.

Dependendo do tamanho do público-alvo, convém adicionar etapas a ou remover etapas deste procedimento.

## <a name="publish-an-office-add-in-via-centralized-deployment"></a>Publicar um suplemento por meio da Implantação Centralizada

Antes de começar, confirme se a sua organização atende a todos os requisitos para usar a Implantação Centralizada, conforme descrito em [Determinar se a Implantação Centralizada de suplementos funciona para sua organização do Office 365](https://support.office.com/pt-BR/article/Determine-if-Centralized-Deployment-of-add-ins-works-for-your-Office-365-organization-B4527D49-4073-4B43-8274-31B7A3166F92).

Se sua organização atender aos requisitos, conclua as etapas a seguir para publicar um suplemento do Office por meio da Implantação Centralizada:

1. Entre no Office 365 com uma conta corporativa ou de estudante.
2. Selecione o ícone do inicializador de aplicativos no canto superior esquerdo e escolha **Administrador**.
3. No menu de navegação, escolha **Configurações** > **Serviços e suplementos**.
4. Se você vir uma mensagem na parte superior da página anunciando o novo Centro de administração do Office 365, escolha a mensagem para ir para a Visualização do Centro de administração (consulte [Sobre o Centro de administração do Office 365](https://support.office.com/en-ie/article/About-the-Office-365-admin-center-758befc4-0888-4009-9f14-0d147402fd23)).
5. Escolha **Carregar um Suplemento** na parte superior da página. 
6. Escolha uma das opções a seguir na página **Implantação Centralizada**:

    - **Desejo adicionar um Suplemento da Office Store.**
    - **Tenho o arquivo de manifesto (.xml) neste dispositivo.** Para esta opção, escolha **Navegar** para localizar o arquivo de manifesto (.xml) que você deseja usar.
    - **Tenho uma URL para o arquivo de manifesto.** Para esta opção, digite a URL do manifesto no campo fornecido.

    ![Caixa de diálogo de novo suplemento no Centro de administração do Office 365](../../images/b3abd42f-63d8-4a5f-8893-d1ae38f4e9b2.png)

7.  Escolha **Avançar**.

8.  Se tiver selecionado a opção para adicionar um suplemento da Office Store, escolha o suplemento. Observe que você pode exibir suplementos disponíveis por meio das categorias **Sugeridos para você**, **Classificação** ou **Nome**. Você só pode adicionar suplementos gratuitos da Office Store; atualmente não é possível adicionar suplementos pagos.

    >**Observação:** Com a opção da Office Store, as atualizações e os aprimoramentos do suplemento serão disponibilizadas automaticamente para usuários sem necessidade de intervenção.

    ![Caixa de diálogo Selecionar um suplemento no Centro de administração do Office 365](../../images/2a8de1f4-03b0-4ab6-aa99-4451ee30a64c.png)

9. O suplemento já está habilitado. Na página para o suplemento, o status é **Ativo**, como o mostrado para o suplemento Bloco do Power BI na captura de tela abaixo. Na seção **Quem tem acesso**, escolha **Editar** para atribuir o suplemento para usuários e/ou grupos.

    ![A página do suplemento Bloco do Power BI no Centro de administração do Office 365](../../images/0faa60e8-1e71-4ed1-bbc1-5a2f85ebf981.png)

10. Na **página Editar quem tem acesso**, escolha **Todos** ou **Usuários/grupos específicos**. Use a caixa Pesquisar para encontrar usuários e/ou grupos para quem você quer implantar o suplemento.

    ![Página Editar quem tem acesso no Centro de administração do Office 365](../../images/46571963-5938-4c7d-b60e-a3ad06758ddf.png)

    >**Observação:** para suplementos de SSO (logon único), os usuários e grupos atribuídos também serão compartilhados com suplementos que compartilham a mesma ID de Aplicativo do Azure. Todas as alterações nas atribuições do usuário também se aplicarão a esses suplementos. Os suplementos relacionados serão mostrados nessa página. Apenas em suplementos de SSO, essa página exibirá a lista de permissões do Microsoft Graph exigida pelo suplemento.

11. Depois de terminar, escolha **Salvar**, revise as configurações do suplemento e escolha **Fechar**. Você verá o suplemento juntamente com outros aplicativos no Office 365.
    >**Observação:** quando um administrador escolhe **Salvar**, o consentimento é fornecido para todos os usuários. 

    ![lista de aplicativos no Centro de administração do Office 365](../../images/71bfd837-20bc-4517-9513-33fc70147669.png)

>**Dica:** Quando você implanta um novo suplemento para usuários e/ou grupos em sua organização, envie um email descrevendo quando e como usar o suplemento e incluindo links para conteúdo relevante da Ajuda, perguntas frequentes ou outros recursos de suporte.

## <a name="considerations-when-granting-access-to-an-add-in"></a>Considerações ao conceder acesso a um suplemento

Os administradores podem atribuir um suplemento a todos na organização ou a usuários e/ou grupos específicos de dentro da organização. A lista a seguir descreve as implicações de cada opção:

- **Todos**: Como o nome sugere, essa opção atribui o suplemento a todos os usuários no locatário. Use essa opção com cautela e apenas para suplementos que sejam realmente universais para sua organização.

- **Usuários**: Se você atribuir um suplemento a usuários individuais, será necessário atualizar as configurações da Central de implantação para o suplemento sempre que quiser atribuí-lo a outros usuários. Da mesma forma, será necessário atualizar as configurações da Central de implantação para o suplemento sempre que você quiser remover o acesso do usuário ao suplemento.

- **Grupos**: Se você atribuir um suplemento a um grupo, os usuários adicionados ao grupo serão atribuídos automaticamente ao suplemento. Da mesma forma, quando um usuário é removido de um grupo, ele automaticamente perde o acesso ao suplemento. Em ambos os casos, nenhuma ação adicional é necessária por parte do administrador do Office 365.

Em geral, para facilitar a manutenção, recomendamos atribuir suplementos usando grupos sempre que possível. No entanto, em situações em que você deseja restringir o acesso do suplemento a um número muito pequeno de usuários, pode ser mais prático atribuir o suplemento a usuários específicos. 

## <a name="add-in-states"></a>Estados de suplementos

A tabela a seguir descreve os estados diferentes de um suplemento.

|Estado|Como o estado ocorre|Impacto|
|-----|--------------------|------|
|**Ativo**|O administrador carregou o suplemento e o atribuiu a usuários e/ou grupos.|Os usuários e/ou grupos atribuídos ao suplemento o veem nos clientes do Office relevantes.|
|**Desativado**|O administrador desativou o suplemento.|Os usuários e/ou grupos atribuídos ao suplemento já não têm acesso a ele. Se o estado do suplemento for alterado de **Desativado** para **Ativo**, os usuários e os grupos recuperarão o acesso a ele.|
|**Excluído**|O administrador excluiu o suplemento.|Os usuários e/ou grupos atribuídos ao suplemento já não têm acesso a ele.|

## <a name="updating-office-add-ins-that-are-published-via-centralized-deployment"></a>Atualizar suplementos do Office que são publicados por meio de Implantação Centralizada

Depois de um suplemento do Office ter sido publicado por meio de Implantação Centralizada, as alterações feitas ao aplicativo Web do suplemento automaticamente estarão disponíveis para todos os usuários assim que as alterações forem implementadas no aplicativo Web. As alterações feitas a um [arquivo de manifesto XML](../overview/add-in-manifests.md) de um suplemento, por exemplo, para atualizar o ícone, texto ou comandos do suplemento, ocorrem da seguinte maneira:

- **Suplemento de linha de negócios**: Se um administrador tiver carregado explicitamente um arquivo de manifesto durante a implementação da Implantação Centralizada por meio do Centro de administração do Office 365, o administrador deverá carregar um novo arquivo de manifesto que contém as alterações desejadas. Depois que o arquivo de manifesto atualizado tiver sido carregado, na próxima vez que os aplicativos relevantes do Office iniciarem, o suplemento será atualizado.

- **Office Store: suplemento**: Se um administrador tiver selecionado um suplemento da Office Store durante a implementação da Implantação Centralizada por meio do Centro de administração do Office 365 e as atualizações de suplementos na Office Store, o suplemento será atualizado posteriormente por meio da Implantação Centralizada. Na próxima vez que os aplicativos relevantes do Office iniciarem, o suplemento será atualizado.

## <a name="end-user-experience-with-add-ins"></a>Experiência do usuário final com suplementos

Depois que um suplemento tiver sido publicado por meio de Implantação Centralizada, os usuários finais podem começar a usá-lo em qualquer plataforma que o suplemento suporte. 

Se o suplemento tiver suporte para comandos, eles serão exibidos na Faixa de Opções do Office a todos os usuários para os quais o suplemento for implantado. No exemplo a seguir, o comando **Pesquisar Citação** aparece na faixa de opções para o suplemento **Citações**. 

![A captura de tela mostra uma seção da faixa de opções do Office com o comando Pesquisar Citação realçado no suplemento Citações](../../images/553b0c0a-65e9-4746-b3b0-8c1b81715a86.png)

Caso contrário, os usuários podem adicioná-lo ao aplicativo do Office da seguinte maneira:

1.  Nos aplicativos Word 2016, Excel 2016 ou PowerPoint 2016, escolha **Inserir** > **Meus Suplementos**.
2.  Escolha **Administrador Gerenciado**, na janela do suplemento.
3.  Escolha o suplemento e escolha **Adicionar**. 

    ![A captura de tela mostra a guia Administração Gerenciada da página Suplementos do Office de um aplicativo do Office. O suplemento Citações é exibido na guia.](../../images/fd36ba81-9882-40f0-9fce-74f991aa97d5.png)
