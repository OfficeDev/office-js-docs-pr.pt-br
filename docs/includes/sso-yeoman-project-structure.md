### <a name="configuration"></a>Configuração

Os seguintes arquivos especificam definições de configuração para o suplemento.

- O arquivo **./manifest.xml** no diretório raiz do projeto define as configurações e os recursos do suplemento.

- O **./. ENV** arquivo no diretório raiz do projeto define constantes que são usadas pelo projeto do suplemento.

### <a name="task-pane"></a>Painel de tarefas 

Os seguintes arquivos definem a interface do usuário e a funcionalidade do painel de tarefas do suplemento.

- O arquivo **./src/taskpane/taskpane.html** contém a marcação HTML do painel de tarefas.

- O arquivo **./src/taskpane/taskpane.css** contém o CSS que é aplicado ao conteúdo no painel de tarefas.

- Em um projeto JavaScript, o arquivo **./src/TaskPane/TaskPane.js** contém código para inicializar o suplemento. Em um projeto TypeScript, o arquivo **./src/TaskPane/TaskPane.TS** contém código para inicializar o suplemento e também o código que usa a biblioteca JavaScript do Office para adicionar os dados do Microsoft Graph ao documento do Office.

### <a name="authentication"></a>Autenticação

Os seguintes arquivos facilitam o processo de SSO e gravam dados no documento do Office.

- Em um projeto JavaScript, o arquivo **./src/Helpers/documentHelper.js** contém código que usa a biblioteca JavaScript do Office para adicionar os dados do Microsoft Graph ao documento do Office. Não há nenhum arquivo em um projeto TypeScript; o código que usa a biblioteca JavaScript do Office para adicionar os dados do Microsoft Graph ao documento do Office existe em **./src/TaskPane/TaskPane.TS** em vez disso.

- O arquivo **./src/Helpers/fallbackauthdialog.html** é a página sem interface do usuário que carrega o JavaScript para a estratégia de autenticação de fallback.

- O arquivo **./src/Helpers/fallbackauthdialog.js** contém o JavaScript para a estratégia de autenticação de fallback que entra no usuário com o MSAL. js.

- O arquivo **./src/Helpers/fallbackauthhelper.js** contém o painel de tarefas JavaScript que invoca a estratégia de autenticação de fallback em cenários em que a autenticação SSO não é suportada.

- O arquivo **./src/helpers/ssoauthhelper.js** contém a chamada JavaScript à API de SSO, `getAccessToken`, recebe o token de inicialização, inicia a troca do token de inicialização por um token de acesso ao Microsoft Graph e chama o Microsoft Graph para obter os dados.