---
title: Autorizar o Microsoft Graph com SSO
description: Saiba como os usuários de um Office suplemento podem usar o SSO (logon único) para buscar dados do Microsoft Graph.
ms.date: 06/10/2022
ms.localizationpriority: medium
ms.openlocfilehash: 4c7bfc51e67755c2a50875f11d3a5477bd5885a4
ms.sourcegitcommit: 4f19f645c6c1e85b16014a342e5058989fe9a3d2
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 06/15/2022
ms.locfileid: "66090940"
---
# <a name="authorize-to-microsoft-graph-with-sso"></a>Autorizar o Microsoft Graph com SSO

Os usuários entrarão no Office usando sua conta microsoft pessoal ou sua conta Microsoft 365 Education ou corporativa. A melhor maneira de um Suplemento do Office receber acesso autorizado ao [Microsoft Graph](https://developer.microsoft.com/graph/docs) é usar as credenciais de logon do Office do usuário. Isso permite a eles acessar seus dados do Microsoft Graph sem precisar entrar novamente.

## <a name="add-in-architecture-for-sso-and-microsoft-graph"></a>Arquitetura de suplemento para SSO e Microsoft Graph

Além de hospedar as páginas e o JavaScript do aplicativo web, o suplemento também deve hospedar, ao mesmo tempo o [nome de domínio totalmente qualificado](/windows/desktop/DNS/f-gly#_dns_fully_qualified_domain_name_fqdn__gly), uma ou mais APIs web que obterá um token de acesso ao Microsoft Graph e fará solicitações a ele.

O manifesto do suplemento contém um elemento **WebApplicationInfo** que fornece informações importantes de registro de aplicativo do Azure para Office, incluindo as permissões para o Microsoft Graph que o suplemento requer.

### <a name="how-it-works-at-runtime"></a>Como ele funciona em tempo de execução

O diagrama a seguir mostra as etapas envolvidas para entrar e acessar o Microsoft Graph. Todo o processo usa tokens de acesso OAuth 2.0 e JWT.

:::image type="content" source="../images/sso-access-to-microsoft-graph.svg" alt-text="Diagrama mostrando o processo de SSO." border="false":::

1. O código do lado do cliente do suplemento chama a API Office.js [getAccessToken](/javascript/api/office-runtime/officeruntime.auth#office-runtime-officeruntime-auth-getaccesstoken-member(1)). Isso instrui o Office host a obter um token de acesso para o suplemento.

    Se o usuário não estiver conectado, o host Office em conjunto com o plataforma de identidade da Microsoft fornecerá a interface do usuário para o usuário entrar e consentir.

2. O Office host solicita um token de acesso do plataforma de identidade da Microsoft.
3. O plataforma de identidade da Microsoft retorna o token *de acesso A* ao Office host. O token *de acesso A* fornece acesso apenas às próprias APIs do lado do servidor do suplemento. Ele não fornece acesso ao Microsoft Graph.
4. O Office host retorna o token *de acesso A* para o código do lado do cliente do suplemento. Agora, o código do lado do cliente pode fazer chamadas autenticadas para as APIs do lado do servidor.
5. O código do lado do cliente faz uma solicitação HTTP para uma API Web no lado do servidor que requer autenticação. Ele inclui o token de acesso *A como* prova de autorização. O código do lado do servidor valida o token de *acesso A*.
6. O código do lado do servidor usa o OBO (fluxo On-Behalf-Of) do OAuth 2.0 para solicitar um novo token de acesso com permissões para o Microsoft Graph.
7. O plataforma de identidade da Microsoft retorna o novo token de acesso *B* com permissões para o Microsoft Graph (e um token de atualização, se o suplemento solicitar *offline_access permissão).* Opcionalmente, o servidor pode armazenar em cache o token *de acesso B*.
8. O código do lado do servidor faz uma solicitação a um microsoft API do Graph e inclui o token de *acesso B* com permissões para o Microsoft Graph.
9. O Microsoft Graph retorna dados de volta para o código do lado do servidor.
10. O código do lado do servidor retorna os dados para o código do lado do cliente.

Em solicitações subsequentes, o código do cliente sempre passará o token *de acesso A* ao fazer chamadas autenticadas para o código do lado do servidor. O código do lado do servidor pode armazenar em cache o token *B* para que ele não precise solicitá-lo novamente em futuras chamadas à API.

## <a name="develop-an-sso-add-in-that-accesses-microsoft-graph"></a>Desenvolver um suplemento SSO que acessa o Microsoft Graph

Você desenvolve um suplemento que acessa o Microsoft Graph como faria com qualquer outro aplicativo que usa o SSO. Para obter uma descrição completa, [consulte Habilitar logon único para Office Suplementos](../develop/sso-in-office-add-ins.md). A diferença é que é obrigatório que o suplemento tenha uma API Web do lado do servidor.

Dependendo do seu idioma e da estrutura, podem estar disponíveis bibliotecas que simplificarão o código do lado do servidor que você precisa escrever. O código deve fazer o seguinte:

* Valide o token *de acesso A* sempre que ele for passado do código do lado do cliente. Para saber mais, confira [Validar o token de acesso](sso-in-office-add-ins.md#pass-the-access-token-to-server-side-code).
* Inicie o OBO (fluxo On-Behalf-Of) do OAuth 2.0 com uma chamada para o plataforma de identidade da Microsoft que inclui o token de acesso, alguns metadados sobre o usuário e as credenciais do suplemento (sua ID e segredo). Para obter mais informações sobre o fluxo OBO, consulte plataforma de identidade da Microsoft e o fluxo [On-Behalf-Of do OAuth 2.0](/azure/active-directory/develop/v2-oauth2-on-behalf-of-flow).
* Opcionalmente, após a conclusão do fluxo, armazene em cache o token de acesso *retornado B* com permissões para o Microsoft Graph. Faça isso se o suplemento fizer mais de uma chamada para o Microsoft Graph. Para obter mais informações, consulte [Adquirir e armazenar tokens em cache usando a MSAL (Biblioteca](/azure/active-directory/develop/msal-acquire-cache-tokens) de Autenticação da Microsoft)
* Crie um ou mais métodos de API Web que obtêm dados do Microsoft Graph passando *o token de* acesso B (possivelmente armazenado em cache) para o Microsoft Graph.

Para obter exemplos detalhados passo a passo de cenários, confira:

* [Criar um Suplemento do Office com Node.js que usa logon único](create-sso-office-add-ins-nodejs.md)
* [Criar um Suplemento do Office com ASP.NET que usa logon único](create-sso-office-add-ins-aspnet.md)
* [Cenário: implementar o logon único no serviço em um suplemento do Outlook](../outlook/implement-sso-in-outlook-add-in.md)

## <a name="distributing-sso-enabled-add-ins-in-microsoft-appsource"></a>Distribuindo suplementos habilitados para SSO no Microsoft AppSource

Quando um Microsoft 365 adquire um suplemento do [AppSource](https://appsource.microsoft.com), o administrador pode redistribuí-lo por meio de Aplicativos Integrados [](/microsoft-365/admin/manage/test-and-deploy-microsoft-365-apps) e conceder consentimento do administrador ao suplemento para acessar os escopos do Microsoft Graph. No entanto, também é possível que o usuário final adquira o suplemento diretamente do AppSource. Nesse caso, o usuário deve conceder consentimento ao suplemento. Isso pode criar um possível problema de desempenho para o qual fornecemos uma solução.

`allowConsentPrompt` `getAccessToken`Se o código passar a opção na chamada de , `OfficeRuntime.auth.getAccessToken( { allowConsentPrompt: true } );`como , Office poderá solicitar o consentimento do usuário se o plataforma de identidade da Microsoft relatar Office que o consentimento ainda não foi concedido ao suplemento. No entanto, por motivos de segurança, Office pode solicitar que o usuário consenta com o escopo Graph `profile` Microsoft. *Office não pode solicitar consentimento para outros escopos Graph Microsoft*, nem mesmo `User.Read`. Isso significa que, se o usuário conceder consentimento no prompt, Office retornará um token de acesso. Mas a tentativa de trocar o token de acesso por um novo token de acesso com escopos adicionais do Microsoft Graph falha com o erro AADSTS65001, o que significa que o consentimento (para escopos do Microsoft Graph) não foi concedido.

> [!NOTE]
> A solicitação de consentimento ainda `{ allowConsentPrompt: true }` poderá falhar mesmo para o `profile` escopo se o administrador tiver desativado o consentimento do usuário final. Para obter mais informações, consulte [Configurar como os usuários finais consentem com aplicativos usando Azure Active Directory](/azure/active-directory/manage-apps/configure-user-consent).

Seu código pode e deve lidar com esse erro voltando para um sistema alternativo de autenticação, o que solicita ao usuário consentimento para escopos Graph Microsoft. Para obter exemplos de código, consulte Criar um suplemento do [Node.js Office](create-sso-office-add-ins-nodejs.md) que usa logon único e criar um suplemento do [ASP.NET Office](create-sso-office-add-ins-aspnet.md) que usa logon único e os exemplos aos quais eles se vinculam. Todo o processo requer várias viagens de ida e volta para o plataforma de identidade da Microsoft. Para evitar essa penalidade de desempenho, inclua `forMSGraphAccess` a opção na chamada de `getAccessToken`; por exemplo, `OfficeRuntime.auth.getAccessToken( { forMSGraphAccess: true } )`. Isso sinaliza para Office que seu suplemento precisa de escopos Graph Microsoft. Office solicitará que o plataforma de identidade da Microsoft verifique se o consentimento para os escopos do Microsoft Graph já foi concedido ao suplemento. Se tiver, o token de acesso será retornado. Caso contrário, a chamada retorna `getAccessToken` o erro 13012. Seu código pode lidar com esse erro voltando para um sistema alternativo de autenticação imediatamente, sem fazer uma tentativa condenada de trocar tokens com o plataforma de identidade da Microsoft.

Como prática recomendada, sempre passe `forMSGraphAccess` `getAccessToken` para quando o suplemento será distribuído no AppSource e precisará de escopos Graph Microsoft.

## <a name="details-on-sso-with-an-outlook-add-in"></a>Detalhes sobre o SSO com um Outlook suplemento

Se você desenvolver um suplemento do Outlook que usa o SSO e o realizar o sideload dele para teste, o Office sempre retornará o erro  13012 `forMSGraphAccess` `getAccessToken` quando for passado, mesmo que o consentimento do administrador tenha sido concedido. Por esse motivo, você deve comentar a opção `forMSGraphAccess` **ao desenvolver** um Outlook suplemento. Remova a marca de comentário da opção ao implantar para produção. O falso 13012 só acontece quando você está descarregando em Outlook.

Para Outlook suplementos, certifique-se de habilitar a Autenticação Moderna para Microsoft 365 locatário. Confira mais informações sobre como fazer isso em [Exchange Online: como habilitar seu locatário para autenticação moderna](https://social.technet.microsoft.com/wiki/contents/articles/32711.exchange-online-how-to-enable-your-tenant-for-modern-authentication.aspx).

## <a name="see-also"></a>Confira também

* [Token OAuth2 Exchange](https://tools.ietf.org/html/draft-ietf-oauth-token-exchange-02)
* [Plataforma de identidade da Microsoft e Fluxo On-Behalf-Of do OAuth 2.0](/azure/active-directory/develop/v2-oauth2-on-behalf-of-flow)
* [Conjuntos de requisitos de IdentityAPI](/javascript/api/requirement-sets/common/identity-api-requirement-sets)
