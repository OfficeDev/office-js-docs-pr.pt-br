---
title: Autorizar o Microsoft Graph com SSO
description: Saiba como os usuários de um Office add-in podem usar o SSO (login único) para buscar dados do Microsoft Graph.
ms.date: 01/25/2022
ms.localizationpriority: medium
---

# <a name="authorize-to-microsoft-graph-with-sso"></a>Autorizar o Microsoft Graph com SSO

Os usuários entram no Office (plataformas online, de dispositivos móveis e de área de trabalho) usando contas pessoais da Microsoft, contas corporativas ou do Microsoft 365 Education. A melhor maneira de um Suplemento do Office receber acesso autorizado ao [Microsoft Graph](https://developer.microsoft.com/graph/docs) é usar as credenciais de logon do Office do usuário. Isso permite a eles acessar seus dados do Microsoft Graph sem precisar entrar novamente.

## <a name="add-in-architecture-for-sso-and-microsoft-graph"></a>Arquitetura de suplemento para SSO e Microsoft Graph

Além de hospedar as páginas e o JavaScript do aplicativo web, o suplemento também deve hospedar, ao mesmo tempo o [nome de domínio totalmente qualificado](/windows/desktop/DNS/f-gly#_dns_fully_qualified_domain_name_fqdn__gly), uma ou mais APIs web que obterá um token de acesso ao Microsoft Graph e fará solicitações a ele.

O manifesto do add-in contém um elemento **WebApplicationInfo** que fornece informações importantes de registro do aplicativo do Azure para Office, incluindo as permissões para o Microsoft Graph que o complemento exige.

### <a name="how-it-works-at-runtime"></a>Como ele funciona em tempo de execução

O diagrama a seguir mostra as etapas envolvidas para entrar e acessar o Microsoft Graph. Todo o processo usa tokens de acesso OAuth 2.0 e JWT.

:::image type="content" source="../images/sso-access-to-microsoft-graph.svg" alt-text="Diagrama mostrando o processo de SSO." border="false":::

1. O código do lado do cliente do complemento chama a API Office.js [getAccessToken](/javascript/api/office-runtime/officeruntime.auth#office-runtime-officeruntime-auth-getaccesstoken-member(1)). Isso diz ao Office host para obter um token de acesso para o complemento.

    Se o usuário não estiver Office, o host Office em conjunto com o plataforma de identidade da Microsoft fornece interface do usuário para o usuário entrar e consentir.

2. O Office host solicita um token de acesso do plataforma de identidade da Microsoft.
3. O plataforma de identidade da Microsoft retorna o token *de acesso A* para o Office host. O token *de acesso A* fornece acesso apenas às APIs do lado do servidor do próprio add-in. Ele não fornece acesso ao Microsoft Graph.
4. O Office host retorna o token *de acesso A* ao código do lado do cliente do complemento. Agora, o código do lado do cliente pode fazer chamadas autenticadas para as APIs do lado do servidor.
5. O código do lado do cliente faz uma solicitação HTTP para uma API da Web no lado do servidor que requer autenticação. Ele inclui o token de acesso *A* como prova de autorização. O código do lado do servidor valida o token de acesso *A*.
6. O código do lado do servidor usa o OAuth 2.0 On-Behalf-Of flow (OBO) para solicitar um novo token de acesso com permissões para o Microsoft Graph.
7. O plataforma de identidade da Microsoft retorna o novo token de acesso *B* com permissões para o Microsoft Graph (e um token de atualização, se o complemento solicitar *offline_access* permissão). O servidor pode, opcionalmente, armazenar em cache o token de acesso *B*.
8. O código do lado do servidor faz uma solicitação para uma API do Microsoft Graph e inclui o token de *acesso B* com permissões para o Microsoft Graph.
9. O Microsoft Graph retorna dados ao código do lado do servidor.
10. O código do lado do servidor retorna os dados para o código do lado do cliente.

Em solicitações subsequentes, o código do cliente sempre passará o token *de acesso A* ao fazer chamadas autenticadas para o código do servidor. O código do lado do servidor pode armazenar em cache o token *B* para que ele não precise solicitá-lo novamente em futuras chamadas de API.

## <a name="develop-an-sso-add-in-that-accesses-microsoft-graph"></a>Desenvolver um suplemento SSO que acessa o Microsoft Graph

Você desenvolve um complemento que acessa o Microsoft Graph como faria com qualquer outro aplicativo que use o SSO. Para uma descrição completa, consulte [Enable single sign-on for Office Add-ins](../develop/sso-in-office-add-ins.md). A diferença é que é obrigatório que o complemento tenha uma API Web do lado do servidor.

Dependendo do seu idioma e da estrutura, podem estar disponíveis bibliotecas que simplificarão o código do lado do servidor que você precisa escrever. O código deve fazer o seguinte:

* Valide o token *de acesso A* sempre que ele for passado do código do lado do cliente. Para saber mais, confira [Validar o token de acesso](sso-in-office-add-ins.md#pass-the-access-token-to-server-side-code).
* Inicie o OAuth 2.0 On-Behalf-Of flow (OBO) com uma chamada para o plataforma de identidade da Microsoft que inclui o token de acesso, alguns metadados sobre o usuário e as credenciais do complemento (sua ID e segredo). Para obter mais informações sobre o fluxo OBO, [consulte plataforma de identidade da Microsoft e OAuth 2.0 On-Behalf-Of flow](/azure/active-directory/develop/v2-oauth2-on-behalf-of-flow).
* Opcionalmente, após a conclusão do fluxo, cache o token de acesso *retornado B* com permissões para o Microsoft Graph. Faça isso se o suplemento fizer mais de uma chamada para o Microsoft Graph. Para obter mais informações, consulte [Adquirir e armazenar tokens de cache usando a Biblioteca de Autenticação da Microsoft (MSAL)](/azure/active-directory/develop/msal-acquire-cache-tokens)
* Crie um ou mais métodos de API Web que recebam dados do Microsoft Graph passando o token de acesso (possivelmente armazenado em cache) *B* para o Microsoft Graph.

Para obter exemplos detalhados passo a passo de cenários, confira:

* [Criar um Suplemento do Office com Node.js que usa logon único](create-sso-office-add-ins-nodejs.md)
* [Criar um Suplemento do Office com ASP.NET que usa logon único](create-sso-office-add-ins-aspnet.md)
* [Cenário: implementar o logon único no serviço em um suplemento do Outlook](../outlook/implement-sso-in-outlook-add-in.md)

## <a name="distributing-sso-enabled-add-ins-in-microsoft-appsource"></a>Distribuição de complementos habilitados para SSO no Microsoft AppSource

Quando um administrador Microsoft 365 adquire um complemento do [AppSource](https://appsource.microsoft.com), o administrador pode redistribui-lo por meio de Aplicativos [](/microsoft-365/admin/manage/test-and-deploy-microsoft-365-apps) Integrados e conceder consentimento ao administrador para que o add-in acesse escopos do Microsoft Graph. No entanto, também é possível que o usuário final adquira o complemento diretamente do AppSource, nesse caso, o usuário deve conceder consentimento ao complemento. Isso pode criar um possível problema de desempenho para o qual fornecemos uma solução.

Se seu `allowConsentPrompt` `getAccessToken`código passar a opção na chamada de , `OfficeRuntime.auth.getAccessToken( { allowConsentPrompt: true } );`como , Office pode solicitar o consentimento do usuário se o plataforma de identidade da Microsoft relata para Office esse consentimento ainda não foi concedido ao add-in. No entanto, por motivos de segurança, Office só pode solicitar que o usuário consenta com o escopo Graph `profile` Microsoft. *Office não pode solicitar o consentimento para outros escopos Graph Microsoft*, nem mesmo `User.Read`. Isso significa que, se o usuário conceder consentimento no prompt, Office retornará um token de acesso. Mas a tentativa de trocar o token de acesso por um novo token de acesso com escopos adicionais do Microsoft Graph falha com o erro AADSTS65001, o que significa que o consentimento (para escopos do Microsoft Graph) não foi concedido.

> [!NOTE]
> A solicitação de consentimento com `{ allowConsentPrompt: true }` ainda poderá falhar mesmo para o `profile` escopo se o administrador tiver desligado o consentimento do usuário final. Para obter mais informações, [consulte Configure how end-users consent to applications using Azure Active Directory](/azure/active-directory/manage-apps/configure-user-consent).

Seu código pode e deve lidar com esse erro voltando para um sistema alternativo de autenticação, que solicita ao usuário consentimento para escopos Graph Microsoft. Para exemplos de código, consulte [Create a Node.js Office Add-in that uses single sign-on](create-sso-office-add-ins-nodejs.md) and [Create an ASP.NET Office Add-in that uses single sign-on](create-sso-office-add-ins-aspnet.md) and the samples they link to. Todo o processo requer várias idas e voltas para o plataforma de identidade da Microsoft. Para evitar essa penalidade de desempenho, inclua a `forMSGraphAccess` opção na chamada de `getAccessToken`; por exemplo, `OfficeRuntime.auth.getAccessToken( { forMSGraphAccess: true } )`. Isso sinaliza para Office que seu add-in precisa de escopos Graph Microsoft. Office solicitará que o plataforma de identidade da Microsoft verifique se o consentimento para os escopos do microsoft Graph já foi concedido ao add-in. Se tiver, o token de acesso será retornado. Se não tiver, a chamada de `getAccessToken` retorna o erro 13012. Seu código pode lidar com esse erro voltando para um sistema alternativo de autenticação imediatamente, sem fazer uma tentativa de troca de tokens com o plataforma de identidade da Microsoft.

Como prática prática, sempre passe `forMSGraphAccess` `getAccessToken` para quando o seu add-in será distribuído no AppSource e precisa de escopos Graph Microsoft.

## <a name="details-on-sso-with-an-outlook-add-in"></a>Detalhes sobre o SSO com um Outlook de dados

Se você desenvolver um Outlook que usa o SSO e fazer sideload dele para testes, o Office sempre retornará o erro 13012  `forMSGraphAccess` `getAccessToken` quando for passado para, mesmo que o consentimento do administrador tenha sido concedido. Por esse motivo, você deve comentar a `forMSGraphAccess` opção **ao desenvolver** um Outlook de usuário. Certifique-se de descompactar a opção ao implantar para produção. O falso 13012 só acontece quando você está fazendo sideload no Outlook.

Para Outlook de Outlook, certifique-se de habilitar a Autenticação Moderna para o Microsoft 365 de autenticação. Confira mais informações sobre como fazer isso em [Exchange Online: como habilitar seu locatário para autenticação moderna](https://social.technet.microsoft.com/wiki/contents/articles/32711.exchange-online-how-to-enable-your-tenant-for-modern-authentication.aspx).

## <a name="see-also"></a>Confira também

* [Token OAuth2 Exchange](https://tools.ietf.org/html/draft-ietf-oauth-token-exchange-02)
* [Plataforma de identidade da Microsoft e Fluxo On-Behalf-Of do OAuth 2.0](/azure/active-directory/develop/v2-oauth2-on-behalf-of-flow)
* [Conjuntos de requisitos IdentityAPI](../reference/requirement-sets/identity-api-requirement-sets.md)
