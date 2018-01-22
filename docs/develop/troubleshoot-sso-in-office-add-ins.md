# <a name="troubleshoot-error-messages-for-single-sign-on-sso"></a>Solucionar problemas de mensagens de erro no logon único (SSO)

Este artigo fornece algumas orientações sobre como solucionar problemas com o logon único (SSO) nos suplementos do Office e como fazer com que seu suplemento habilitado para SSO lide de forma robusta com os erros ou condições especiais.

## <a name="debugging-tools"></a>Ferramentas de depuração

Recomendamos fortemente que você use uma ferramenta que possa interceptar e exibir as solicitações HTTP a partir de seu serviço Web do suplemento, além de respostas para ele, quando você estiver desenvolvendo. Duas das ferramentas mais populares são: 

- [Fiddler](http://www.telerik.com/fiddler): Gratuita ([Documentação](http://docs.telerik.com/fiddler/configure-fiddler/tasks/configurefiddler))
- [Charles](https://www.charlesproxy.com/): Gratuita por 30 dias. ([Documentação](https://www.charlesproxy.com/documentation/))

Ao desenvolver sua API de serviço, você também pode tentar:

- [Postman](http://www.getpostman.com/postman): Gratuita ([Documentação](https://www.getpostman.com/docs/))

## <a name="causes-and-handling-of-errors-from-getaccesstokenasync"></a>Causas e tratamento dos erros do getAccessTokenAsync

### <a name="13000"></a>13000

A API [getAccessTokenAsync](http://dev.office.com/reference/add-ins/shared/office.context.auth.getAccessTokenAsync) não é compatível pelo suplemento ou pela versão do Office. 

- A versão do Office não é compatível com o SSO. A versão necessária é o Office 2016, versão 1710, build 8629.nnnn ou posterior (a versão de assinatura do Office 365, às vezes chamada de "Clique para Executar"). Você talvez precise ser um participante do programa Office Insider para obter essa versão. Para obter mais informações, confira a página [Seja um Office Insider](https://products.office.com/en-us/office-insider?tab=tab-1). 
- O manifesto do suplemento está sem a seção [WebApplicationInfo](http://dev.office.com/reference/add-ins/manifest/webapplicationinfo) adequada.

### <a name="13001"></a>13001

O usuário não iniciou sessão no Office. Seu código deve chamar novamente o método `getAccessTokenAsync` e passar a opção `forceAddAccount: true` no parâmetro [opções](../../reference/shared/office.context.auth.getAccessTokenAsync.md#parameters). 

### <a name="13002"></a>13002

O usuário cancelou o login ou o consentimento. 
- Se o seu suplemento fornece funções que não exigem que o usuário esteja conectado (ou que tenha concedido o consentimento), seu código deve capturar esse erro e permitir que o suplemento permaneça em execução.
- Se o suplemento exigir um usuário conectado que concedeu consentimento, seu código deve solicitar ao usuário que repita a operação, mas não mais do que uma vez. 

### <a name="13003"></a>13003

Tipo de Usuário não suportado. O usuário não iniciou sessão no Office com uma conta Microsoft válida ou uma conta comercial ou escolar. Isso pode acontecer se o Office funcionar com uma conta de domínio no local, por exemplo. Seu código deve solicitar ao usuário que faça login no Office.

### <a name="13004"></a>13004

Recurso inválido. O manifesto do suplemento não foi configurado corretamente. Atualize o manifesto. Para obter mais informações, consulte [Validar e solucionar problemas com seu manifesto](troubleshoot-manifest.md).

### <a name="13005"></a>13005

Concessão inválida. Isso geralmente significa que o Office não foi pré-autorizado para o serviço Web do suplemento. Para obter mais informações, consulte [Criar o aplicativo de serviço](../../docs/develop/sso-in-office-add-ins.md#create-the-service-application) e [Registrar o suplemento com o ponto de extremidade v2.0 do Azure AD](../../docs/develop/create-sso-office-add-ins-aspnet.md#register-the-add-in-with-azure-ad-v2-0-endpoint) (ASP.NET) ou [Registrar o suplemento com o ponto de extremidade v2.0 do Azure AD](../../docs/develop/create-sso-office-add-ins-nodejs.md#register-the-add-in-with-azure-ad-v2-0-endpoint) (Nó JS). Isso também pode acontecer se o usuário não concedeu as permissões de aplicativo de serviço para o seu `profile`.

### <a name="13006"></a>13006

Erro do cliente. Seu código deve sugerir que o usuário saia e reinicie o Office.

### <a name="13007"></a>13007

O host do Office não conseguiu obter um token de acesso ao serviço Web do suplemento.
- Certifique-se de que seu registro de suplemento e seu manifesto de suplemento especifiquem as permissões `openid` e `profile`. Para obter mais informações, consulte [Registrar o suplemento com o ponto de extremidade v2.0 do Azure AD](../../docs/develop/create-sso-office-add-ins-aspnet.md#register-the-add-in-with-azure-ad-v2-0-endpoint) (ASP.NET) ou [Registrar o suplemento com o ponto de extremidade v2.0 do Azure AD](../../docs/develop/create-sso-office-add-ins-nodejs.md#register-the-add-in-with-azure-ad-v2-0-endpoint) (Nó JS), e [Configurar o suplemento](../../docs/develop/create-sso-office-add-ins-aspnet.md#configure-the-add-in) (ASP.NET) ou [Configurar o suplemento](../../docs/develop/create-sso-office-add-ins-nodejs.md#configure-the-add-in) (Nó JS).
- Seu código pode sugerir que o usuário tente novamente a operação.

### <a name="13008"></a>13008

O usuário desencadeou uma operação que chama o `getAccessTokenAsync` antes de uma chamada anterior do `getAccessTokenAsync` concluída. Seu código deve solicitar ao usuário que repita a operação após a operação anterior ter sido concluída.

## <a name="errors-on-the-server-side-from-azure-active-directory"></a>Erros no lado do servidor do Azure Active Directory

### <a name="conditional-access--multifactor-authentication-errors"></a>Erros no Acesso condicional/Autenticação multifatorial
 
Em certas configurações de identidade no AAD e no Office 365, é possível que alguns recursos que são acessíveis com o Microsoft Graph exijam autenticação multifator (MFA), mesmo quando o locatário do Office 365 do usuário não exija. Quando o AAD recebe uma solicitação de um token para o recurso protegido por MFA, através do fluxo Em Nome De, ele retorna ao serviço Web do seu suplemento uma mensagem JSON que contém uma propriedade `claims`. A propriedade de reivindicações tem informações sobre quais outros fatores de autenticação são necessários. 

Seu código do lado do servidor deve testar esta mensagem e transmitir o valor das reivindicações ao seu código do lado do cliente. Você precisa dessa informação no cliente porque o Office processa a autenticação para os suplementos de SSO. A mensagem para o cliente pode ser um erro (como `500 Server Error` ou `401 Unauthorized`) ou estar no corpo de uma resposta de sucesso (como `200 OK`). Em ambos os casos, o retorno de chamada (falha ou sucesso) da chamada AJAX do lado do cliente do seu código para a API da Web do seu suplemento deve testar essa resposta. Se o valor das solicitações tiver sido retransmitido, seu código deve chamar novamente o `getAccessTokenAsync` e passar a opção `authChallenge: CLAIMS-STRING-HERE` no parâmetro [opções](../../reference/shared/office.context.auth.getAccessTokenAsync.md#parameters). Quando o AAD vir essa string, ele solicita ao usuário o(s) fator(es) adicional(ais) e, em seguida, retorna um novo token de acesso que será aceito no fluxo Em Nome De.

Temos algumas amostras para ilustrar este tratamento MFA: 

- [Suplemento do Office ASPNET SSO](https://github.com/OfficeDev/Office-Add-in-ASPNET-SSO): A biblioteca MSAL que esta amostra usa expõe a mensagem MFA do AAD como uma exceção. O código retransmite isso ao cliente como uma resposta `500 Server Error`. No script do lado do cliente, o retorno de chamada `fail` da chamada AJAX chama novamente o `getAccessTokenAsync` com a opção `authChallenge`. Veja especificamente os arquivos [ValuesController.cs](https://github.com/OfficeDev/Office-Add-in-ASPNET-SSO/blob/master/Complete/Office-Add-in-ASPNET-SSO-WebAPI/Controllers/ValuesController.cs) e [Home.js](https://github.com/OfficeDev/Office-Add-in-ASPNET-SSO/blob/master/Complete/Office-Add-in-ASPNET-SSO-WebAPI/Scripts/Home.js).
- [Suplemento do Office NodeJS SSO](https://github.com/OfficeDev/Office-Add-in-NodeJS-SSO): A mensagem MFA do AAD é enviada ao cliente como uma resposta de sucesso. No script do lado do cliente, o retorno de chamada `done` da chamada AJAX chama novamente o `getAccessTokenAsync` com a opção `authChallenge`. Veja especificamente os arquivos [auth.ts](https://github.com/OfficeDev/Office-Add-in-NodeJS-SSO/blob/master/Completed/src/auth.ts) e [program.js](https://github.com/OfficeDev/Office-Add-in-NodeJS-SSO/blob/master/Completed/public/program.js).

### <a name="consent-missing-errors"></a>Erros de falta de consentimento

Se o AAD não tiver um registro de que o consentimento (para o recurso Microsoft Graph) foi concedido ao suplemento pelo usuário (ou administrador do locatário), o AAD enviará uma mensagem de erro ao seu serviço Web. Seu código deve dizer ao cliente (no corpo de uma resposta `403 Forbidden`, por exemplo) para chamar novamente o `getAccessTokenAsync` com a opção `forceConsent: true`.

### <a name="invalid-or-missing-scope-permission-errors"></a>Erros de escopo (permissão) inválidos ou ausentes

- Seu código do lado do servidor deve enviar uma resposta `403 Forbidden` ao cliente que deve apresentar uma mensagem amigável ao usuário. Se possível, registre o erro no console ou grave-o em um registro.
- Certifique-se de que sua seção de [Escopos](http://dev.office.com/reference/add-ins/manifest/scopes) do manifesto do suplemento especifique todas as permissões necessárias. E certifique-se de que seu registro do serviço Web do suplemento especifique as mesmas permissões. Verifique também os erros de ortografia. Para obter mais informações, consulte [Registrar o suplemento com o ponto de extremidade v2.0 do Azure AD](../../docs/develop/create-sso-office-add-ins-aspnet.md#register-the-add-in-with-azure-ad-v2-0-endpoint) (ASP.NET) ou [Registrar o suplemento com o ponto de extremidade v2.0 do Azure AD](../../docs/develop/create-sso-office-add-ins-nodejs.md#register-the-add-in-with-azure-ad-v2-0-endpoint) (Nó JS), e [Configurar o suplemento](../../docs/develop/create-sso-office-add-ins-aspnet.md#configure-the-add-in) (ASP.NET) ou [Configurar o suplemento](../../docs/develop/create-sso-office-add-ins-nodejs.md#configure-the-add-in) (Nó JS).

### <a name="expired-or-invalid-token-errors-when-calling-microsoft-graph"></a>Erros de token expirados ou inválidos ao chamar o Microsoft Graph

Algumas bibliotecas de autenticação e autorização, incluindo o MSAL, evitam erros de token expirados usando um token de atualização em cache sempre que necessário. Você também pode codificar seu próprio sistema de cache de token. Para uma amostra que faz isso, consulte [Suplemento do Office NodeJS SSO](https://github.com/OfficeDev/Office-Add-in-NodeJS-SSO), especialmente o arquivo [auth.ts](https://github.com/OfficeDev/Office-Add-in-NodeJS-SSO/blob/master/Completed/src/auth.ts).

Mas se você receber um token expirado ou um erro de token inválido, seu código deve dizer ao cliente (no corpo de uma resposta `401 Unauthorized`, por exemplo) para chamar novamente o `getAccessTokenAsync` e repetir a chamada para o ponto de extremidade de sua API da Web do suplemento, que repetirá o fluxo Em Nome De para obter um novo token para o Microsoft Graph. 

### <a name="invalid-token-error-when-calling-microsoft-graph"></a>Erro de token inválido ao chamar novamente o Microsoft Graph

Lide com esse erro da mesma forma que um erro de token expirado. Veja a seção anterior.

### <a name="invalid-audience-error"></a>Erro de audiência inválido

Seu código do lado do servidor deve enviar uma resposta `403 Forbidden` ao cliente que deve apresentar uma mensagem amigável ao usuário e, possivelmente, também registrar o erro no console ou gravá-lo em um registro.

Para obter mais informações sobre a adição de suporte de vários locatários para validação de token, consulte a [Amostra de vários locatários do Azure](https://github.com/Azure-Samples/active-directory-dotnet-webapp-webapi-multitenant-openidconnect).
