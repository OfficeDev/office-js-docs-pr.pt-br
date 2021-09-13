### <a name="configuration"></a>Configuração

Os arquivos a seguir especificam as configurações do complemento.

- O arquivo **./manifest.xml** no diretório raiz do projeto define as configurações e os recursos do suplemento.

- O **./. O arquivo ENV** no diretório raiz do projeto define constantes que são usadas pelo projeto do complemento.

### <a name="task-pane"></a>Painel de tarefas 

Os arquivos a seguir definem a interface do usuário e a funcionalidade do painel de tarefas do complemento.

- O arquivo **./src/taskpane/taskpane.html** contém a marcação HTML do painel de tarefas.

- O arquivo **./src/taskpane/taskpane.css** contém o CSS que é aplicado ao conteúdo no painel de tarefas.

- Em um projeto JavaScript, o **arquivo ./src/taskpane/taskpane.js** contém código para inicializar o add-in. Em um projeto TypeScript, o **arquivo ./src/taskpane/taskpane.ts** contém código para inicializar o add-in e também o código que usa a biblioteca de API JavaScript do Office para adicionar os dados do Microsoft Graph ao documento Office.

### <a name="authentication"></a>Autenticação

Os arquivos a seguir facilitam o processo de SSO e escrevem dados no Office documento.

- Em um projeto JavaScript, o **arquivo ./src/helpers/documentHelper.js** contém código que usa Office biblioteca de API JavaScript do Office para adicionar os dados do Microsoft Graph ao documento Office. Não existe esse arquivo em um projeto TypeScript; o código que usa Office biblioteca da API JavaScript para adicionar os dados do Microsoft Graph ao documento Office existe em **./src/taskpane/taskpane.ts** em vez disso.

- O **arquivo ./src/helpers/fallbackauthdialog.html** é a página sem interface do usuário que carrega o JavaScript para a estratégia de autenticação de fallback.

- O **arquivo ./src/helpers/fallbackauthdialog.js** contém o JavaScript para a estratégia de autenticação de fallback que faz a assinatura no usuário com msal.js.

- O **arquivo ./src/helpers/fallbackauthhelper.js** contém o painel de tarefas JavaScript que invoca a estratégia de autenticação de fallback em cenários em que a autenticação SSO não é suportada.

- O arquivo **./src/helpers/ssoauthhelper.js** contém a chamada JavaScript à API de SSO, `getAccessToken`, recebe o token de inicialização, inicia a troca do token de inicialização por um token de acesso ao Microsoft Graph e chama o Microsoft Graph para obter os dados.