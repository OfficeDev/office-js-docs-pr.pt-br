---
title: Autenticação e autorização com a API da caixa de diálogo do Office
description: Aprenda a usar a API da caixa de diálogo do Office para permitir que os usuários entrem no Google, no Facebook, no Microsoft 365 e em outros serviços protegidos pela Plataforma de Identidade da Microsoft.
ms.date: 07/22/2021
ms.localizationpriority: high
ms.openlocfilehash: aa4ce5b74752623e10b61082d6f9becc1a26b713
ms.sourcegitcommit: 45f7482d5adcb779a9672669360ca4d8d5c85207
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 01/19/2022
ms.locfileid: "62074186"
---
# <a name="authenticate-and-authorize-with-the-office-dialog-api"></a>Autenticação e autorização com a API da caixa de diálogo do Office

Várias autoridades de identidade, também chamadas de Serviços de Token Seguro (STS), impedem que a página de logon seja aberta em um IFrame. Isso inclui o Google, o Facebook e os serviços protegidos pela Plataforma de Identidade da Microsoft (antigo Azure AD V 2.0), como uma conta da Microsoft, uma conta corporativa ou de estudante do Microsoft 365 ou outra conta comum. Isso cria um problema para os suplementos do Office, porque quando o suplemento é executado no **Office na Web**, o painel de tarefas é um IFrame. Os usuários de um suplemento só podem fazer logon em um desses serviços se o suplemento puder abrir uma instância do navegador completamente separada. Isso porque o Office fornece a [API da Caixa de Diálogo](dialog-api-in-office-add-ins.md), especificamente o método [displayDialogAsync](/javascript/api/office/office.ui).

> [!NOTE]
> Esse artigo presume que você esteja familiarizado com o [Uso da API da Caixa de Diálogo do Office nos suplementos do Office.](dialog-api-in-office-add-ins.md).

A caixa de diálogo aberta com essa API tem as seguintes características.

- [Não é restrita](https://en.wikipedia.org/wiki/Dialog_box).
- É uma instância do navegador completamente separada do painel de tarefas, ou seja:
  - Tem o seu próprio ambiente de tempo de execução do JavaScript, objeto de janela e variáveis globais.
  - Não há nenhum ambiente de execução compartilhado com o painel de tarefas.
  - Não compartilha o mesmo armazenamento de sessão (a propriedade [Window.sessionStorage](https://developer.mozilla.org/docs/Web/API/Window/sessionStorage)) como o painel de tarefas.
- A primeira página aberta na caixa de diálogo deve estar hospedada no mesmo domínio que o painel de tarefas, incluindo o protocolo, os subdomínios e a porta, se houver.
- A caixa de diálogo pode enviar informações de volta para o painel de tarefas usando o método [messageParent](/javascript/api/office/office.ui#messageParent_message__messageOptions_). Recomendamos que esse método seja chamado somente de uma página hospedada no mesmo domínio que o painel de tarefas, incluindo protocolo, subdomínios e porta. Caso contrário, haverá complicações em como você chama o método e processa a mensagem. Para obter mais informações, [mensagens entre domínios para o runtime do host](dialog-api-in-office-add-ins.md#cross-domain-messaging-to-the-host-runtime).


Por padrão, a caixa de diálogo é aberta em um controle de exibição da Web totalmente novo, não em um iframe. Isso garante que ele possa abrir a página de logon de um provedor de identidade. Como será mostrado neste artigo, as características da caixa de diálogo têm implicações sobre como você usa as bibliotecas de autenticação ou autorização, como a MSAL e o Passport.

> [!NOTE]
> Para configurar a caixa de diálogo para abrir em um iframe flutuante: passe a opção `displayInIframe: true` na chamada do `displayDialogAsync`. *Não* faça isso quando estiver usando a API da Caixa de Diálogo do Office para logon.

## <a name="authentication-flow-with-the-office-dialog-box"></a>Fluxo de autenticação com a caixa de diálogo do Office

A seguir está um fluxo de autenticação típico.

![Diagrama mostrando a relação entre o painel de tarefas e os processos do navegador de caixa de diálogo.](../images/taskpane-dialog-processes.gif)

1. A primeira página que é aberta na caixa de diálogo é uma página local (ou outro recurso) que está hospedada no domínio do suplemento; ou seja, o mesmo domínio da janela do painel de tarefas. Essa página pode ter uma única interface de usuário que informa "Aguarde. Estamos redirecionando você para a página onde poderá entrar no *NOME-DO-PROVEDOR*." O código nessa página constrói a URL da página de entrada do provedor de identidade usando as informações que são transmitidas para a caixa de diálogo, conforme descrito em [Transmitir informações para a caixa de diálogo](dialog-api-in-office-add-ins.md#pass-information-to-the-dialog-box) ou é codificado em um arquivo de configuração do suplemento, como um arquivo web.config.
2. A janela da caixa de diálogo redireciona então para a página de entrada. A URL inclui um parâmetro de consulta que informa o provedor de identidade para redirecionar a janela da caixa de diálogo a uma página específica após o usuário entrar. Nesse artigo, chamaremos essa página de **redirectPage.html**. Nesta página, os resultados da tentativa de entrada podem ser passados para o painel de tarefas com uma chamada de `messageParent`. *Recomendamos que esta seja uma página no mesmo domínio que a janela do host*.
3. O serviço do provedor de identidade processa a solicitação GET recebida da janela da caixa de diálogo. Se o usuário já estiver conectado, ele imediatamente redirecionará a janela para **redirectPage.html** e incluirá os dados do usuário como um parâmetro de consulta. Se o usuário ainda não tiver entrado, a página de entrada do provedor aparecerá na janela para que o usuário possa entrar. Para a maioria dos provedores, se o usuário não consegue entrar com êxito, o provedor mostra uma página de erro na janela da caixa de diálogo e não redireciona para **redirectPage.html**. O usuário precisa fechar a janela selecionando o **X** no canto. Se o usuário entrar com êxito, a janela de diálogo será redirecionada para **redirectPage.html** e os dados do usuário serão incluídos como um parâmetro de consulta.
4. Quando a página **redirectPage.html** é aberta, ela chama a `messageParent` para relatar o êxito ou a falha na página do painel de tarefas e opcionalmente também pode informar os dados do usuário ou os dados de erro. Outras mensagens possíveis incluem passar um token de acesso ou informar ao painel de tarefas que o token está no armazenamento.
5. O evento `DialogMessageReceived` é acionado na página do painel de tarefas, seu manipulador fecha a janela da caixa de diálogo e assim a mensagem pode ser processada.

#### <a name="support-multiple-identity-providers"></a>Prestar suporte a vários provedores de identidade

Se o seu suplemento oferece ao usuário diversas opções de provedores, como a conta da Microsoft, o Google ou o Facebook, você precisa de uma primeira página local (confira a seção anterior) que forneça uma interface de usuário para a escolha de um provedor. A escolha do provedor acionará a construção do URL de entrada e o redirecionamento para ele.

#### <a name="authorization-of-the-add-in-to-an-external-resource"></a>Autorização do suplemento para um recurso externo

Na Web moderna, os usuários e aplicativos Web são entidades de segurança. O aplicativo tem sua própria identidade e permissões para recursos online, como o Microsoft 365, Google Plus, Facebook ou LinkedIn. O aplicativo é registrado no provedor de recursos antes da implantação. O registro inclui:

- Uma lista das permissões que o aplicativo precisa.
- Uma URL para a qual o serviço do recurso deve retornar um token de acesso quando o aplicativo acessa o serviço.  

Quando um usuário invoca uma função no aplicativo que acessa os dados do usuário no serviço do recurso, ele é solicitado a entrar no serviço e a conceder ao aplicativo as permissões necessárias para os recursos do usuário. Em seguida, o serviço redireciona a janela de entrada para a URL previamente registrada e transmite o token de acesso. O aplicativo usa o token de acesso para acessar os recursos do usuário.

Você pode usar as APIs de Caixa de Diálogo do Office para gerenciar esse processo usando um fluxo semelhante àquele descrito para os usuários entrarem. As únicas diferenças são:

- Se o usuário ainda não tiver concedido ao aplicativo as permissões necessárias, será solicitado a fazê-lo na caixa de diálogo após entrar.
- A janela da caixa de diálogo envia o token de acesso à janela do host usando `messageParent` para enviar o token de acesso em formato de cadeia de caracteres ou armazenando o token de acesso em um local onde a janela do host poderá recuperá-lo (e usando `messageParent` para informar à janela do host que o token está disponível). O token tem um limite de tempo, mas enquanto durar, a janela do host poder usá-lo para acessar recursos do usuário de forma direta, sem outras solicitações.

Alguns suplementos de exemplo de autenticação que usam a API da Caixa de Diálogo do Office para essa finalidade estão listados em [Amostras](#samples).

## <a name="using-authentication-libraries-with-the-dialog-box"></a>Usar bibliotecas de autenticação pela caixa de diálogo

O fato de a Caixa de Diálogo do Office e o painel de tarefas serem executados em navegadores diferentes e no tempo de execução do JavaScript, as instâncias significam que você deve usar muitas bibliotecas de autenticação/autorização de maneira diferente de como elas são usadas quando a autenticação e a autorização ocorrem na mesma janela. As seções a seguir descrevem as principais maneiras pelas quais, geralmente, você não pode usar essas bibliotecas e a maneira que você *pode* usá-las.

### <a name="you-usually-cannot-use-the-librarys-internal-cache-to-store-tokens"></a>Geralmente, você não pode usar o cache interno da biblioteca para armazenar tokens

Normalmente, as bibliotecas relacionadas à autenticação fornecem um cache na memória para armazenar o token de acesso. Se chamadas subsequentes para o provedor de recursos (por exemplo, Google, Microsoft Graph, Facebook, etc.) forem feitas, a biblioteca primeiro verificará se o token no cache está expirado. Caso não tenha expirado, a biblioteca retornará o token em cache, em vez de retornar ao STS para obter um novo token. No entanto, esse padrão não pode ser usado em Suplementos do Office. Uma vez que o logon ocorre na instância do navegador da caixa de diálogo do Office, o cache do token estará nessa instância.

Estritamente relacionado a isso está o fato de que uma biblioteca normalmente fornece métodos interativos e "silenciosos" para obter um token. Quando for possível fazer tanto a autenticação quanto as chamadas de dados ao recurso na mesma instância do navegador, o código chamará o método silencioso para obter um token imediatamente antes do código adicionar o token à chamada de dados. O método silencioso procurará por um token não expirado no cache e o retornará, caso haja um. Caso contrário, o método silencioso chamará o método interativo que será redirecionado para o logon do STS. Após a conclusão do logon, o método interativo retorna o token e o armazena na memória. No entanto, quando a API da Caixa de Diálogo do Office está sendo usada, as chamadas de dados do recurso, que chamam o método silencioso, estão na instância do navegador do painel de tarefas. O cache de token da biblioteca não existe nessa instância.

Como alternativa, a instância do navegador da Caixa de Diálogo do suplemento pode chamar diretamente o método interativo da biblioteca. Quando esse método retorna um token, o código deve armazenar explicitamente o token em algum lugar onde a instância do navegador do painel de tarefas pode recuperá-lo, como o Armazenamento Local\* ou um banco de dados do lado do servidor. Outra opção é passar o token para o painel de tarefas com o método `messageParent`. Essa alternativa só é possível se o método interativo armazenar o token de acesso em um local onde o código possa lê-lo. Às vezes, o método interativo de uma biblioteca é projetado para armazenar o token em uma propriedade particular de um objeto que está inacessível ao código.

> [!NOTE]
> \* Há um bug que afetará sua estratégia de tratamento de tokens. Se o suplemento estiver sendo executado no **Office na Web** nos navegadores Safari ou Microsoft Edge, o painel de tarefas e a caixa de diálogo não compartilharão o mesmo Armazenamento Local, portanto, ele não poderá ser usado para a comunicação entre eles.

### <a name="you-usually-cannot-use-the-librarys-auth-context-object"></a>Geralmente, você não pode usar o objeto "contexto de autenticação" da biblioteca

Frequentemente, uma biblioteca relacionada à autenticação possui um método que obtém tanto um token de forma interativa, como também cria um objeto de "contexto de autenticação" retornado pelo método. O token é uma propriedade do objeto (possivelmente particular e inacessível diretamente do código). Esse objeto tem os métodos que recebem os dados do recurso. Esses métodos incluem o token nas Solicitações HTTP feitas ao provedor de recursos (por exemplo, Google, Microsoft Graph, Facebook, etc.).

Esses objetos de contexto de autenticação e os métodos que os criam não podem ser usados nos Suplementos do Office. Como o logon ocorre na instância do navegador da caixa de diálogo do Office, o objeto teria que ser criado lá. Mas as chamadas de dados do recurso estão na instância do navegador do painel de tarefas e não há como enviar o objeto de uma instância para outra. Por exemplo, não é possível passar o objeto pelo `messageParent` porque `messageParent` só pode passar valores de cadeia de caracteres. Um objeto do JavaScript com métodos não pode ser transformado em cadeia de caracteres de maneira confiável.

### <a name="how-you-can-use-libraries-with-the-office-dialog-api"></a>Como usar as bibliotecas através da API da Caixa de Diálogo do Office

Além dos objetos monolíticos de "contexto de autenticação", a maioria das bibliotecas fornecem APIs em um nível inferior de abstração que permite que o código crie objetos auxiliares menos monolíticos. Por exemplo, [MSAL.NET](https://github.com/AzureAD/microsoft-authentication-library-for-dotnet/wiki#conceptual-documentation) v. 3.x.x tem uma API para construir uma URL de logon e outra API que constrói um objeto AuthResult que contém um token de acesso em uma propriedade que pode ser acessada pelo código. Para obter exemplos de MSAL.NET em um Suplemento do Office, confira: [ASP.NET Microsoft Graph no Suplemento do Office](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth/Office-Add-in-Microsoft-Graph-ASPNET) e [ASP.NET Microsoft Graph no Suplemento do Outlook](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth/Outlook-Add-in-Microsoft-Graph-ASPNET). Para ver um exemplo de como usar o [msal.js](https://github.com/AzureAD/microsoft-authentication-library-for-js) em um suplemento, confira [Microsoft Graph React no Suplemento do Office](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth/Office-Add-in-Microsoft-Graph-React).

Para saber mais sobre as bibliotecas de autenticação e autorização, confira [Microsoft Graph: bibliotecas recomendadas](authorize-to-microsoft-graph-without-sso.md#recommended-libraries-and-samples) e [Outros serviços externos: bibliotecas](auth-external-add-ins.md#libraries).

## <a name="samples"></a>Exemplos

- [ASP.NET Microsoft Graph no Suplemento do Office](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth/Office-Add-in-Microsoft-Graph-ASPNET): um suplemento com base em ASP.NET (Excel, Word ou PowerPoint) que usa a biblioteca MSAL.NET e o Fluxo de Código de Autorização para efetuar logon, e obter um token de acesso para dados do Microsoft Graph.
- [ASP.NET Microsoft Graph no Suplemento do Outlook](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth/Outlook-Add-in-Microsoft-Graph-ASPNET): semelhante a exibida acima, mas o aplicativo do Office sendo o Outlook.
- [Microsoft Graph React no Suplemento do Office](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth/Office-Add-in-Microsoft-Graph-React): um suplemento com base em NodeJS (Excel, Word ou PowerPoint) que usa a biblioteca msal.js e o Fluxo Implícito para efetuar logon, e obter um token de acesso para dados do Microsoft Graph.

## <a name="see-also"></a>Conferir também

- [Autorizar serviços externos no Suplemento do Office](auth-external-add-ins.md)
- [Usar a API da Caixa de Diálogo do Office nos suplementos do Office](dialog-api-in-office-add-ins.md)
