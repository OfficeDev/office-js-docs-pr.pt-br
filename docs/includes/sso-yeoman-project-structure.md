### <a name="configuration"></a>Configuração

Os arquivos a seguir especificam definições de configuração para o suplemento.

- O arquivo **./manifest.xml** no diretório raiz do projeto define as configurações e os recursos do suplemento.

- O arquivo **./.ENV** no diretório raiz do projeto define as constantes que são usadas pelo projeto de suplemento.

### <a name="task-pane"></a>Painel de tarefas

Os arquivos a seguir definem a interface do usuário e a funcionalidade do painel de tarefas do suplemento.

- O arquivo **./src/taskpane/taskpane.html** contém a marcação HTML do painel de tarefas.

- O arquivo **./src/taskpane/taskpane.css** contém o CSS que é aplicado ao conteúdo no painel de tarefas.

- Em um projeto JavaScript, o arquivo **./src/taskpane/taskpane.js** contém código para inicializar o suplemento. Em um projeto TypeScript, o arquivo **./src/taskpane/taskpane.ts** contém código para inicializar o suplemento e também código que usa a biblioteca de API JavaScript do Office para adicionar os dados do Microsoft Graph ao documento do Office.

### <a name="authentication"></a>Autenticação

Os arquivos a seguir facilitam o processo de logon único e gravam dados no documento do Office.

- Em um projeto JavaScript, o arquivo **./src/helpers/documentHelper.js** contém código que usa a biblioteca de API JavaScript do Office para adicionar os dados do Microsoft Graph ao documento do Office. Não existe esse arquivo em um projeto TypeScript; o código que usa a biblioteca da API JavaScript do Office para adicionar os dados do Microsoft Graph ao documento do Office existe em **./src/taskpane/taskpane.ts**.

- O arquivo **./src/helpers/fallbackauthdialog.html** é a página sem interface do usuário que carrega o JavaScript para a estratégia de autenticação de fallback.

- O arquivo **./src/helpers/fallbackauthdialog.js** contém o JavaScript para a estratégia de autenticação de fallback que conecta o usuário com msal.js.

- O arquivo **./src/helpers/fallbackauthhelper.js** contém o JavaScript do painel de tarefas que invoca a estratégia de autenticação de fallback em cenários em que a autenticação de logon único não tem suporte.

- O arquivo **./src/helpers/ssoauthhelper.js** contém a chamada JavaScript para a API de logon único, `getAccessToken`, recebe o token de acesso, inicia a troca do token de acesso para um novo token de acesso com permissões para o Microsoft Graph e chama o Microsoft Graph para os dados.