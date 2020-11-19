---
title: Autorizar o Microsoft Graph com SSO
description: Saiba como os usuários de um suplemento do Office podem usar o logon único (SSO) para buscar dados do Microsoft Graph.
ms.date: 07/30/2020
localization_priority: Normal
ms.openlocfilehash: e87c86b5302bde8122485b837759fa327251c656
ms.sourcegitcommit: ceb8dd66f3fb9c963fce8446c2f6c65ead56fbc1
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 11/18/2020
ms.locfileid: "49131910"
---
# <a name="authorize-to-microsoft-graph-with-sso"></a>Autorizar o Microsoft Graph com SSO

Os usuários entram no Office (plataformas online, de dispositivos móveis e de área de trabalho) usando contas pessoais da Microsoft, contas corporativas ou do Microsoft 365 Education. A melhor maneira de um Suplemento do Office receber acesso autorizado ao [Microsoft Graph](https://developer.microsoft.com/graph/docs) é usar as credenciais de logon do Office do usuário. Isso permite a eles acessar seus dados do Microsoft Graph sem precisar entrar novamente.

> [!NOTE]
> A API de Logon Único é compatível com Word, Excel, Outlook e PowerPoint. Confira mais informações sobre os programas para os quais a API de logon único tem suporte no momento em [Conjuntos de requisitos da IdentityAPI](/office/dev/add-ins/reference/requirement-sets/identity-api-requirement-sets).
> Se você estiver trabalhando com um suplemento do Outlook, certifique-se de habilitar a Autenticação Moderna para a locação do Office 365. Confira mais informações sobre como fazer isso em [Exchange Online: como habilitar seu locatário para autenticação moderna](https://social.technet.microsoft.com/wiki/contents/articles/32711.exchange-online-how-to-enable-your-tenant-for-modern-authentication.aspx).

## <a name="add-in-architecture-for-sso-and-microsoft-graph"></a>Arquitetura de suplemento para SSO e Microsoft Graph

Além de hospedar as páginas e o JavaScript do aplicativo web, o suplemento também deve hospedar, ao mesmo tempo o [nome de domínio totalmente qualificado](/windows/desktop/DNS/f-gly#_dns_fully_qualified_domain_name_fqdn__gly), uma ou mais APIs web que obterá um token de acesso ao Microsoft Graph e fará solicitações a ele.

O manifesto do suplemento contém a marcação que especifica como ele está registrado no ponto de extremidade v2.0 do Azure Active Directory (Azure AD) e especifica todas as permissões para o Microsoft Graph que o suplemento precisa.

### <a name="how-it-works-at-runtime"></a>Como ele funciona em tempo de execução

O diagrama a seguir mostra como funciona o processo de entrar e obter acesso ao Microsoft Graph.

![Diagrama mostrando o processo de SSO](../images/sso-access-to-microsoft-graph.png)

1. No suplemento, o JavaScript chama uma nova API do Office.js [getAccessToken](/javascript/api/office-runtime/officeruntime.auth#getaccesstoken-options-). Isso informa ao aplicativo cliente do Office para obter um token de acesso para o suplemento. (De agora em diante, isso se chamará **token de acesso de inicialização** porque é substituído por um segundo token mais tarde durante o processo. Para ver um exemplo de um token de acesso de inicialização decodificado, confira [Token de acesso de exemplo](sso-in-office-add-ins.md#example-access-token).)
2. Se o usuário não estiver conectado, o aplicativo cliente do Office abrirá uma janela pop-up para o usuário entrar.
3. Se essa é a primeira vez que o usuário atual usa seu suplemento, será solicitado que ele dê o consentimento.
4. O aplicativo cliente do Office solicita o **token de acesso de inicialização** do ponto de extremidade v 2.0 do Azure ad para o usuário atual.
5. O Azure AD envia o token de inicialização para o aplicativo cliente do Office.
6. O aplicativo cliente do Office envia o **token de acesso de inicialização** para o suplemento como parte do objeto de resultado retornado pela `getAccessToken` chamada.
7. O JavaScript no suplemento faz uma solicitação HTTP a uma API Web que está hospedada no mesmo domínio totalmente qualificado que o suplemento e inclui o **token de acesso de inicialização** como prova de autorização.
8. O código no lado do servidor valida o **token de acesso de inicialização** de entrada.
9. O código do lado do servidor usa o fluxo "em nome de" (definido em [OAuth2 token Exchange](https://tools.ietf.org/html/draft-ietf-oauth-token-exchange-02) e o [daemon ou aplicativo de servidor para o cenário da API Web do Azure](/azure/active-directory/develop/active-directory-authentication-scenarios)) para obter um token de acesso para o Microsoft Graph no Exchange para o token de acesso de inicialização.
10. O Azure AD retorna o token de acesso de inicialização para o Microsoft Graph (e um token de atualização, se o suplemento solicitar a permissão *offline_access*) para ele próprio.
11. O código do lado do servidor armazena em cache o token de acesso ao Microsoft Graph.
12. O código do lado do servidor faz solicitações ao Microsoft Graph e inclui o token de acesso ao Microsoft Graph.
13. O Microsoft Graph retorna dados para o suplemento, que pode passá-los para a interface do usuário do suplemento.
14. Quando o token de acesso ao Microsoft Graph expira, o código do lado do servidor pode usar o token de atualização para obter um novo token de acesso ao Microsoft Graph.

## <a name="develop-an-sso-add-in-that-accesses-microsoft-graph"></a>Desenvolver um suplemento SSO que acessa o Microsoft Graph

Você desenvolve um suplemento que acessa o Microsoft Graph como faria com qualquer outro suplemento que use SSO. Para obter uma descrição completa, confira [Habilitar o logon único para Suplementos do Office](../develop/sso-in-office-add-ins.md). A diferença é que é obrigatório que o suplemento tenha uma API Web do lado do servidor, e o token de acesso nesse artigo é chamado de "token de acesso de inicialização".

Dependendo do seu idioma e da estrutura, podem estar disponíveis bibliotecas que simplificarão o código do lado do servidor que você precisa escrever. O código deve fazer o seguinte:

* Inicie o fluxo "em nome de" com uma chamada para o ponto de extremidade do Azure AD v 2.0 que inclui o token de acesso de inicialização, alguns metadados sobre o usuário e as credenciais do suplemento (ID e segredo).
* Crie um ou mais métodos de API Web que obtêm dados do Microsoft Graph passando o token de acesso (possivelmente em cache) para o Microsoft Graph.
* Opcionalmente, antes de iniciar o fluxo, valide o token de acesso de inicialização que é recebido do manipulador de token que você criou anteriormente. Para saber mais, confira [Validar o token de acesso](sso-in-office-add-ins.md#validate-the-access-token). 
* Opcionalmente, após concluir o fluxo, armazene em cache o token de acesso retornado no Microsoft Graph. Faça isso se o suplemento fizer mais de uma chamada para o Microsoft Graph. Para mais informações sobre esse fluxo, confira [Azure Active Directory v2.0 e fluxo "On-Behalf-Of" do OAuth 2.0](/azure/active-directory/develop/active-directory-v2-protocols-oauth-on-behalf-of).

> [!NOTE]
> Para exemplos de tokens de acesso decodificados do Microsoft Graph obtidos pelo fluxo "on behalf of", confira [Azure Active Directory v2.0 e fluxo "On-Behalf-Of" do OAuth 2.0](/azure/active-directory/develop/active-directory-v2-protocols-oauth-on-behalf-of).

Para obter exemplos detalhados passo a passo de cenários, confira:

* [Criar um Suplemento do Office com Node.js que usa logon único](create-sso-office-add-ins-nodejs.md)
* [Criar um Suplemento do Office com ASP.NET que usa logon único](create-sso-office-add-ins-aspnet.md)
* [Cenário: implementar o logon único no serviço em um suplemento do Outlook](../outlook/implement-sso-in-outlook-add-in.md)

## <a name="distributing-sso-enabled-add-ins-in-microsoft-appsource"></a>Distribuindo suplementos habilitados para SSO no Microsoft AppSource

Quando um administrador do Microsoft 365 adquire um suplemento do [AppSource](https://appsource.microsoft.com), o administrador pode redistribuí-lo pela [implantação centralizada](../publish/centralized-deployment.md) e conceder consentimento do administrador ao suplemento para acessar escopos do Microsoft Graph. Também é possível, no entanto, que o usuário final adquira o suplemento diretamente de AppSource, caso em que o usuário deve conceder consentimento ao suplemento. Isso pode criar um possível problema de desempenho para o qual fornecemos uma solução.

Se seu código passar a `allowConsentPrompt` opção na chamada de `getAccessToken` , like `OfficeRuntime.auth.getAccessToken( { allowConsentPrompt: true } );` , o Office poderá solicitar o consentimento do usuário se o Azure ad reportar ao Office que o consentimento ainda não foi concedido ao suplemento. No entanto, por motivos de segurança, o Office só pode solicitar que o usuário concorde com o escopo do Azure AD `profile` . *O Office não pode solicitar o consentimento para os escopos do Microsoft Graph*, nem mesmo `User.Read` . Isso significa que, se o usuário conceder consentimento no prompt, o Office retornará um token de inicialização. Mas a tentativa de trocar o token de inicialização para um token de acesso ao Microsoft Graph falhará com o erro AADSTS65001, o que significa que o consentimento (para escopos do Microsoft Graph) não foi concedido.

Seu código pode e deve lidar com esse erro recorrendo a um sistema alternativo de autenticação, o que solicitará o consentimento do usuário aos escopos do Microsoft Graph. (Para obter exemplos de código, confira [criar um Node.js suplemento do Office que usa o logon único](create-sso-office-add-ins-nodejs.md) e [criar um suplemento do Office do ASP.NET que usa o logon único](create-sso-office-add-ins-aspnet.md) e os exemplos aos quais eles se vinculam.) Mas todo o processo requer vários tempos de viagem para o Azure AD. Você pode evitar essa penalidade de desempenho, incluindo a `forMSGraphAccess` opção na chamada de `getAccessToken` ; por exemplo, `OfficeRuntime.auth.getAccessToken( { forMSGraphAccess: true } )` .  Isso indica ao Office que seu suplemento precisa de escopos do Microsoft Graph. O Office solicitará que o Azure AD Verifique se o consentimento para os escopos do Microsoft Graph já foi concedido ao suplemento. Se tiver sido, o token Bootstrap será retornado. Se não tiver, a chamada de retornará o `getAccessToken` erro 13012. Seu código pode lidar com esse erro recorrendo a um sistema alternativo de autenticação imediatamente, sem fazer uma tentativa de Doomed para trocar tokens com o Azure AD.

Como prática recomendada, sempre passe `forMSGraphAccess` para `getAccessToken` quando o suplemento for distribuído no AppSource e precise de escopos do Microsoft Graph.

> [!TIP]
> Se você desenvolver um suplemento do Outlook que usa SSO e Sideload-lo para teste, o Office *sempre* retornará o erro 13012 quando `forMSGraphAccess` for passado para `getAccessToken` mesmo se o consentimento do administrador tiver sido concedido. Por esse motivo, você deve comentar a `forMSGraphAccess` opção **ao desenvolver** um suplemento do Outlook. Certifique-se de remover o comentário da opção ao implantar para produção. O 13012 falso só acontece quando você está Sideload no Outlook.
