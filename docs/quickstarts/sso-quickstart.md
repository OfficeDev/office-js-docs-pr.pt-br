---
title: Início rápido logon único (SSO).
description: Use o gerador Yeoman para construir um Suplemento Office Node.js que utilize o logon único.
ms.date: 09/07/2022
ms.prod: non-product-specific
ms.localizationpriority: high
ms.openlocfilehash: ecbecfd7e475c224451735c7a864f6de2c230d07
ms.sourcegitcommit: cff5d3450f0c02814c1436f94cd1fc1537094051
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 09/30/2022
ms.locfileid: "68239367"
---
# <a name="single-sign-on-sso-quick-start"></a>Início rápido logon único (SSO).

Neste artigo, você usará o gerador Yeoman para Suplementos do Office para criar um Suplemento do Office para Excel, Outlook, Word ou PowerPoint que usa SSO (logon único).

> [!NOTE]
> O modelo de SSO fornecido pelo gerador Yeoman para Suplementos do Office só é executado no localhost e não pode ser implantado. Se você estiver criando um novo Suplemento do Office com SSO para fins de produção, siga as instruções em Criar um suplemento do [Node.js Office](../develop/create-sso-office-add-ins-nodejs.md) que usa logon único.

## <a name="prerequisites"></a>Pré-requisitos

- [Node.js](https://nodejs.org) (a versão mais recente de [LTS](https://nodejs.org/about/releases)).

- A versão mais recente do [Yeoman](https://github.com/yeoman/yo) e do [Yeoman gerador de Suplementos do Office](../develop/yeoman-generator-overview.md). Para instalar essas ferramentas globalmente, execute o seguinte comando por meio do prompt de comando.

    ```command&nbsp;line
    npm install -g yo generator-office
    ```

    [!include[note to update Yeoman generator](../includes/note-yeoman-generator-update.md)]

- Se você estiver usando um Mac e não tiver a CLI do Azure instalada no computador, instale o [Homebrew](https://brew.sh/). O script de configuração do SSO executado durante o início rápido usará o Homebrew para instalar a CLI do Azure e, em seguida, usará a CLI do Azure para configurar o SSO no Azure.

## <a name="create-the-add-in-project"></a>Criar o projeto do suplemento

> [!TIP]
> O gerador Yeoman pode criar um Suplemento do Office habilitado para SSO para Excel, Outlook, Word ou PowerPoint com o tipo de script JavaScript ou TypeScript. As instruções a seguir especificam o `JavaScript` e o `Excel`, mas você deverá escolher o tipo de script e o aplicativo cliente do Office que atendem melhor ao seu cenário.

[!include[Yeoman generator create project guidance](../includes/yo-office-command-guidance.md)]

- **Escolha o tipo de projeto:** `Office Add-in Task Pane project supporting single sign-on (localhost)`
- **Escolha o tipo de script:** `JavaScript`
- **Qual será o nome do suplemento?** `My Office Add-in`
- **Você gostaria de proporcionar suporte para qual aplicativo cliente do Office?** Escolha `Excel`, `Outlook`, `Word`ou `Powerpoint`.

:::image type="content" source="../images/yo-office-sso-excel.png" alt-text="Solicitações e respostas para o gerador Yeoman em uma interface de linha de comando.":::

Depois que você concluir o assistente, o gerador criará o projeto e instalará os componentes Node de suporte.

[!include[Yeoman generator next steps](../includes/yo-office-next-steps.md)]

## <a name="explore-the-project"></a>Explore o projeto

O projeto de suplemento que você criou com o gerador do Yeoman contém um código para um suplemento de painel de tarefas habilitado para SSO.

[!include[project structure for an SSO-enabled add-in created with the Yeoman generator](../includes/sso-yeoman-project-structure.md)]

## <a name="configure-sso"></a>Configure o SSO

Agora que seu projeto de suplemento foi criado e contém o código necessário para facilitar o processo de SSO, conclua as etapas a seguir para configurar o SSO para seu suplemento.

1. Vá para a pasta raiz do projeto.

    ```command&nbsp;line
    cd "My Office Add-in"
    ```

2. Execute o comando a seguir para configurar o SSO do suplemento.

    ```command&nbsp;line
    npm run configure-sso
    ```

    > [!WARNING]
    > Esse comando falhará se o locatário estiver configurado para exigir autenticação de dois fatores. Nesse cenário, você precisará concluir manualmente o registro de aplicativo do Azure e as etapas de configuração de SSO seguindo todas as etapas no tutorial Criar um suplemento do [Office do Node.js](../develop/create-sso-office-add-ins-nodejs.md) que usa o logon único.

3. Uma janela de navegador da Web será exibida e solicitará que você entre no Azure.  Entre no Azure com as suas credenciais de administrador do Microsoft 365. Essas credenciais serão usadas para registrar um novo aplicativo no Azure e definir as configurações necessárias para o SSO.

    > [!NOTE]
    > Se você entrar no Azure usando credenciais de não administrador durante essa etapa, o script `configure-sso` não conseguirá fornecer consentimento de administrador para o suplemento aos usuários da organização. Portanto, o SSO não estará disponível aos usuários do suplemento e eles serão solicitados a entrar.

4. Depois de inserir suas credenciais, feche a janela do navegador e retorne ao prompt de comando. Durante o processo de configuração do SSO, você verá mensagens de status sendo gravadas no console. Conforme descrito nas mensagens do console, os arquivos no projeto do suplemento que o gerador Yeoman criou são atualizados automaticamente com os dados necessários ao processo de SSO.

## <a name="test-your-add-in"></a>Testar seu suplemento

Se você criou um suplemento do Excel, Word ou PowerPoint, conclua as etapas na seção a seguir para experimentá-lo. Se você criou um suplemento do Outlook, conclua as etapas na seção [do Outlook](#outlook) .

### <a name="excel-word-and-powerpoint"></a>Excel, Word e PowerPoint

Conclua as etapas a seguir para testar um suplemento do Excel, Word ou PowerPoint.

1. Quando o processo de configuração do SSO for concluído, execute o seguinte comando para criar o projeto: inicie o servidor Web local e sideload o suplemento no aplicativo cliente do Office selecionado anteriormente.

    [!INCLUDE [alert use https](../includes/alert-use-https.md)]

    ```command&nbsp;line
    npm start
    ```

2. Quando o Excel, o Word ou o PowerPoint for aberto quando você executar o comando anterior, verifique se você está conectado com uma conta de usuário que seja membro da mesma organização do Microsoft 365 que a conta de administrador do Microsoft 365 que você usou para se conectar ao Azure ao configurar o SSO na etapa 3 da [seção](#configure-sso) anterior. Isso estabelecerá as condições apropriadas para que o SSO seja bem-sucedido.

3. No aplicativo cliente do Office, escolha a  guia Página Inicial e, em  seguida, escolha Mostrar Painel de Tarefas para abrir o painel de tarefas do suplemento.

    :::image type="content" source="../images/excel-quickstart-addin-3b.png" alt-text="Botão do suplemento do Excel.":::

4. Na parte inferior do painel de tarefas, escolha o botão **Obter Informações do Meu Perfil de Usuário** para iniciar o processo de SSO.

5. Se uma janela de diálogo for exibida solicitando permissões em nome do suplemento, isso significa que não há suporte ao SSO no seu cenário e, em vez disso, o suplemento voltou para um método alternativo de autenticação do usuário. Isso pode ocorrer quando o administrador de locatários não concedeu consentimento para o suplemento acessar o Microsoft Graph ou quando o usuário não está conectado ao Office com uma conta Microsoft válida ou uma conta Microsoft 365 Education ou Corporativa. Escolha **Aceitar** para continuar.

    ![Captura de tela mostrando o diálogo de permissão solicitada com o botão Aceitar destacado.](../images/sso-permissions-request.png)

    > [!NOTE]
    > Após um usuário aceitar a solicitação de permissões, elas não serão solicitadas novamente no futuro.

6. O suplemento recupera as informações de perfil do usuário conectado e as grava no documento. A imagem a seguir mostra um exemplo de informações de perfil gravadas em uma planilha do Excel.

    ![Captura de tela mostrando informações de perfil do usuário na planilha do Excel.](../images/sso-user-profile-info-excel.png)

### <a name="outlook"></a>Outlook

Execute as etapas a seguir para experimentar um suplemento do Outlook.

1. Quando concluir o processo de configuração de SSO, execute o seguinte comando para criar o projeto e iniciar o servidor Web local.

    [!INCLUDE [alert use https](../includes/alert-use-https.md)]

    ```command&nbsp;line
    npm start
    ```

2. Siga as instruções [Realizar sideload dos suplementos do Outlook para teste](../outlook/sideload-outlook-add-ins-for-testing.md)para realizar o sideload do suplemento do Outlook. Certifique-se de que você está conectado ao Outlook com um usuário que seja membro da mesma organização do Microsoft 365, como a conta de administrador do Microsoft 365 que você usou para se conectar ao Azure, ao configurar o SSO na etapa 3 da [seção anterior](#configure-sso). Isso estabelecerá as condições apropriadas para que o SSO seja bem-sucedido.

3. Escreva uma nova mensagem no Outlook.

4. Na janela redigir mensagem, escolha o botão Mostrar Painel **de Tarefas** para abrir o painel de tarefas do suplemento.

    ![Captura de tela mostrando o botão da faixa de opções do suplemento destacado na janela de composição de mensagem do Outlook.](../images/outlook-sso-ribbon-button.png)

5. Na parte inferior do painel de tarefas, escolha o botão **Obter Informações do Meu Perfil de Usuário** para iniciar o processo de SSO.

6. Se uma janela de diálogo for exibida solicitando permissões em nome do suplemento, isso significa que não há suporte ao SSO no seu cenário e, em vez disso, o suplemento voltou para um método alternativo de autenticação do usuário. Isso pode ocorrer quando o administrador de locatários não concedeu consentimento para o suplemento acessar o Microsoft Graph ou quando o usuário não está conectado ao Office com uma conta Microsoft válida ou uma conta Microsoft 365 Education ou Corporativa. Escolha **Aceitar** para continuar.

    ![Captura de tela da caixa de diálogo de permissões solicitadas com o botão Aceitar destacado.](../images/sso-permissions-request.png)

    > [!NOTE]
    > Após um usuário aceitar a solicitação de permissões, elas não serão solicitadas novamente no futuro.

7. O suplemento recupera as informações de perfil do usuário conectado e as grava no corpo da mensagem do e-mail.

    ![Captura de tela mostrando informações de perfil de usuário na janela de composição de mensagem do Outlook.](../images/sso-user-profile-info-outlook.png)

## <a name="next-steps"></a>Próximas etapas

Parabéns, você criou com êxito um suplemento do painel de tarefas que usa SSO sempre que possível; e usa um método alternativo de autenticação de usuário quando não há suporte ao SSO. Para saber como personalizar seu suplemento para adicionar novas funcionalidades que requerem permissões diferentes, consulte [Personalizar o suplemento habilitado para SSO do Node.js](sso-quickstart-customize.md).

## <a name="see-also"></a>Confira também

- [Habilitar o logon único para Suplementos do Office](../develop/sso-in-office-add-ins.md)
- [Personalizar o suplemento habilitado para SSO do Node.js](sso-quickstart-customize.md).
- [Criar um Suplemento do Office com Node.js que usa logon único](../develop/create-sso-office-add-ins-nodejs.md)
- [Solucionar problemas de mensagens de erro no logon único (SSO)](../develop/troubleshoot-sso-in-office-add-ins.md)
- [Usando o Visual Studio Code para publicar](../publish/publish-add-in-vs-code.md#using-visual-studio-code-to-publish)