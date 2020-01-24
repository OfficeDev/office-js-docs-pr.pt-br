---
title: Use o gerador Yeoman para criar um Suplemento do Office que use SSO (prévia)
description: Use o gerador Yeoman para criar um Suplemento do Office com Node.js que use logon único (prévia).
ms.date: 01/16/2020
ms.prod: non-product-specific
localization_priority: Priority
ms.openlocfilehash: becc0a03a87dcfd5b37b5ab65f45dd6516bf105a
ms.sourcegitcommit: 8bce9c94540ed484d0749f07123dc7c72a6ca126
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 01/22/2020
ms.locfileid: "41265589"
---
# <a name="use-the-yeoman-generator-to-create-an-office-add-in-that-uses-single-sign-on-preview"></a>Use o gerador Yeoman para criar um Suplemento do Office que use logon único (prévia)

Neste artigo, você seguirá pelo processo de uso do gerador Yeoman para criar um Suplemento do Office para Excel, Word ou PowerPoint que usa o logon único (SSO) sempre que possível, e usa um método alternativo de autenticação do usuário quando não há suporte ao SSO.

> [!TIP]
> Antes de tentar concluir o início rápido, revise [Habilitar o logon único para Suplementos do Office](../develop/sso-in-office-add-ins.md) para aprender conceitos básicos sobre o SSO em Suplementos do Office. 
 
O gerador Yeoman simplifica o processo de criação de um suplemento de SSO, automatizando as etapas necessárias para configurar o SSO no Azure e gerando o código necessário para um suplemento usar o SSO. Para um passo a passo detalhado descrevendo como concluir manualmente as etapas que o gerador Yeoman automatiza, confira o tutorial [Criar um Suplemento do Office com Node.js que usa logon único](../develop/create-sso-office-add-ins-nodejs.md).

## <a name="prerequisites"></a>Pré-requisitos

* [Node.js](https://nodejs.org) (a versão mais recente de [LTS](https://nodejs.org/about/releases))

* A versão mais recente do [Yeoman](https://github.com/yeoman/yo) e do [Yeoman gerador de suplementos do Office](https://github.com/OfficeDev/generator-office). Para instalar essas ferramentas globalmente, execute o seguinte comando por meio do prompt de comando:

    ```command&nbsp;line
    npm install -g yo generator-office
    ```

    [!include[note to update Yeoman generator](../includes/note-yeoman-generator-update.md)]

[!include[additional prerequisites](../includes/sso-tutorial-prereqs.md)]

## <a name="create-the-add-in-project"></a>Crie o projeto do suplemento

> [!TIP]
> O gerador Yeoman pode criar um Suplemento do Office habilitado para SSO do Excel, Word ou PowerPoint e pode ser criado com o tipo de script JavaScript ou TypeScript. As instruções a seguir especificam o `JavaScript` e o `Excel`, mas você deverá escolher o tipo de script e o aplicativo cliente do Office que atendem melhor ao seu cenário.

[!include[Yeoman generator create project guidance](../includes/yo-office-command-guidance.md)]

- **Escolha o tipo de projeto:** `Office Add-in Task Pane project supporting single sign-on`
- **Escolha o tipo de script:** `Javascript`
- **Qual será o nome do suplemento?** `My SSO Office Add-in`
- **Você gostaria de proporcionar suporte para qual aplicativo cliente do Office?** `Excel`

![Uma captura de tela dos prompts e respostas do gerador Yeoman](../images/yo-office-sso-excel.png)

Depois que você concluir o assistente, o gerador criará o projeto e instalará os componentes Node de suporte.

[!include[Yeoman generator next steps](../includes/yo-office-next-steps.md)]

## <a name="explore-the-project"></a>Explore o projeto

O projeto de suplemento que você criou com o gerador do Yeoman contém um código para um suplemento de painel de tarefas habilitado para SSO.

- O arquivo **./manifest.xml** no diretório raiz do projeto define as configurações e os recursos do suplemento.

- O arquivo **./src/taskpane/taskpane.html** contém a marcação HTML do painel de tarefas.
- O arquivo **./src/taskpane/taskpane.css** contém o CSS que é aplicado ao conteúdo no painel de tarefas.
- O arquivo **./src/taskpane/taskpane.js** contém o código da API JavaScript do Office que facilita a interação entre o painel de tarefas e o aplicativo host do Office.

- O arquivo **./src/helpers/documentHelper.js** usa a biblioteca JavaScript do Office para adicionar os dados do Microsoft Graph ao documento do Office.
- O arquivo **./src/helpers/fallbackauthdialog.html** é a página sem interface do usuário que carrega o JavaScript do método de autenticação de fallback.
- O arquivo **./src/Helpers/fallbackauthdialog.js** contém o JavaScript do método de autenticação fallback que entra no usuário com o msal.js.
- O arquivo **./src/helpers/fallbackauthhelper.js** contém o painel de tarefas JavaScript que chama o método de autenticação de fallback em cenários em que não há suporte à autenticação SSO.
- O arquivo **./src/helpers/ssoauthhelper.js** contém a chamada JavaScript à API de SSO, `getAccessToken`, recebe o token de inicialização, inicia a troca do token de inicialização por um token de acesso ao Microsoft Graph e chama o Microsoft Graph para obter os dados.

- O arquivo **./ENV** no diretório raiz do projeto define as constantes que são usadas pelo projeto do suplemento.
    > [!NOTE]
    > Algumas das constantes definidas neste arquivo são usadas para facilitar o processo de SSO. Talvez você queira atualizar os valores nesse arquivo para que eles correspondam ao seu cenário específico. Por exemplo, você pode atualizar o arquivo para especificar um escopo diferente, se o seu suplemento exigir algo diferente de `User.Read`.

## <a name="configure-sso"></a>Configure o SSO

Nesse ponto, seu projeto de suplemento foi criado e contém o código necessário para facilitar o processo de SSO. Depois, execute as etapas a seguir para configurar o SSO do seu suplemento.

1. Navegue até a pasta raiz do projeto.

    ```command&nbsp;line
    cd "My SSO Office Add-in"
    ```

2. Execute o comando a seguir para configurar o SSO do suplemento.

    ```command&nbsp;line
    npm run configure-sso
    ```

    > [!WARNING]
    > Esse comando falhará se o locatário estiver configurado para exigir autenticação de dois fatores. Nesse cenário, será necessário concluir manualmente as etapas de configuração do SSO e registro do aplicativo Azure, conforme descrito no tutorial [Criar um Suplemento do Office com Node.js que usa logon único](../develop/create-sso-office-add-ins-nodejs.md).

3. Uma janela de navegador da Web será exibida e solicitará que você entre no Azure.  Entre no Azure com as suas credenciais de administrador do Office 365. Essas credenciais serão usadas para registrar um novo aplicativo no Azure e definir as configurações necessárias para o SSO.

    > [!NOTE]
    > Se você entrar no Azure usando credenciais de não administrador durante essa etapa, o script `configure-sso` não conseguirá fornecer consentimento de administrador para o suplemento aos usuários da organização. Portanto, o SSO não estará disponível aos usuários do suplemento e eles serão solicitados a entrar.

4. Depois de inserir suas credenciais, feche a janela do navegador e retorne ao prompt de comando. Durante o processo de configuração do SSO, você verá mensagens de status sendo gravadas no console. Conforme descrito nas mensagens do console, os arquivos no projeto do suplemento que o gerador Yeoman criou são atualizados automaticamente com os dados necessários ao processo de SSO.

## <a name="try-it-out"></a>Experimente

1. Quando o processo de configuração do SSO for concluído, execute o seguinte comando para criar o projeto: inicie o servidor Web local e sideload o suplemento no aplicativo cliente do Office selecionado anteriormente.

    > [!NOTE]
    > Os Suplementos do Office devem usar HTTPS, e não HTTP, mesmo durante o desenvolvimento. Se for solicitado a instalação de um certificado após executar um dos seguintes comandos, aceite a solicitação para instalar o certificado que o gerador do Yeoman fornecer.

    ```command&nbsp;line
    npm start
    ```

2. No aplicativo cliente do Office que é aberto ao executar o comando anterior (por exemplo, Excel, Word ou PowerPoint), certifique-se de estar conectado com um usuário que seja membro da mesma organização do Office 365, como uma conta de administrador do Office 365 que você usou para se conectar ao Azure, enquanto configura o SSO na etapa 3 da [seção anterior](#configure-sso). Isso estabelecerá as condições apropriadas para que o SSO seja bem-sucedido. 

3. No aplicativo cliente do Office, escolha a guia **Página Inicial** e o botão **Mostrar Painel de Tarefas** na faixa de opções para abrir o painel de tarefas do suplemento. A imagem a seguir mostra esse botão no Excel.

    ![Botão do suplemento do Excel](../images/excel-quickstart-addin-3b.png)

4. Na parte inferior do painel de tarefas, escolha o botão **Obter Informações do Meu Perfil de Usuário** para iniciar o processo de SSO. 

    > [!NOTE] 
    > Se você ainda não tiver entrado no Office, será solicitado a fazê-lo. Conforme descrito anteriormente, será necessário entrar com um usuário que seja membro da mesma organização do Office 365, como a conta de administrador do Office 365 que você usou para se conectar ao Azure, enquanto configura o SSO na etapa 3 da [seção anterior](#configure-sso), se desejar que o SSO seja bem-sucedido.

5. Se uma janela de diálogo for exibida solicitando permissões em nome do suplemento, isso significa que não há suporte ao SSO no seu cenário e, em vez disso, o suplemento voltou para um método alternativo de autenticação do usuário. Isso pode ocorrer quando o administrador do locatário não tiver consentido ao suplemento acesso ao Microsoft Graph, ou quando o usuário não estiver conectado ao Office com uma conta válida da Microsoft ou do Office 365 ("Corporativa ou de Estudante"). Escolha o botão **Aceitar** na janela de diálogo para continuar.

    ![Caixa de diálogo Solicitação de permissões](../images/sso-permissions-request.png)

    > [!NOTE]
    > Após um usuário aceitar a solicitação de permissões, elas não serão solicitadas novamente no futuro.

6. O suplemento recupera as informações de perfil do usuário conectado e as grava no documento. A imagem a seguir mostra um exemplo de informações de perfil gravadas em uma planilha do Excel.

    ![Informações de perfil de usuário na planilha do Excel](../images/sso-user-profile-info-excel.png)

## <a name="next-steps"></a>Próximas etapas

Parabéns, você criou com êxito um suplemento do painel de tarefas que usa SSO sempre que possível; e usa um método alternativo de autenticação de usuário quando não há suporte ao SSO. Para saber mais sobre as etapas de configuração do SSO que o gerador Yeoman concluiu automaticamente e o código que facilita o processo de SSO, confira o tutorial [Criar um Suplemento do Office com Node.js que usa logon único](../develop/create-sso-office-add-ins-nodejs.md).

## <a name="see-also"></a>Confira também

- [Habilitar o logon único para Suplementos do Office](../develop/sso-in-office-add-ins.md)
- [Criar um Suplemento do Office com Node.js que usa logon único](../develop/create-sso-office-add-ins-nodejs.md)
- [Solucionar problemas de mensagens de erro no logon único (SSO)](../develop/troubleshoot-sso-in-office-add-ins.md)