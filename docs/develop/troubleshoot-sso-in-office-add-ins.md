---
title: Solucionar problemas de mensagens de erro no logon único (SSO)
description: ''
ms.date: 12/08/2017
ms.openlocfilehash: 5abf10d8281ea54be9a172c3f45b742fb33991df
ms.sourcegitcommit: c53f05bbd4abdfe1ee2e42fdd4f82b318b363ad7
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 10/12/2018
ms.locfileid: "25506067"
---
# <a name="troubleshoot-error-messages-for-single-sign-on-sso-preview"></a>Solucionar problemas de mensagens de erro no logon único (SSO) (versão prévia)

Este artigo fornece algumas orientações sobre como solucionar problemas com o logon único (SSO) nos suplementos do Office e como fazer com que seu suplemento habilitado para SSO trate de forma robusta os erros ou condições especiais.

> [!NOTE]
> |||UNTRANSLATED_CONTENT_START|||The Single Sign-on API is currently supported in preview for Word, Excel, Outlook, and PowerPoint. For more information about where the Single Sign-on API is currently supported, see [IdentityAPI requirement sets]https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/identity-api-requirement-sets). To use SSO, you must load the beta version of the Office JavaScript Library from https://appsforoffice.microsoft.com/lib/beta/hosted/office.js in the startup HTML page of the add-in. If you are working with an Outlook add-in, be sure to enable Modern Authentication for the Office 365 tenancy. For information about how to do this, see [Exchange Online: How to enable your tenant for modern authentication](https://social.technet.microsoft.com/wiki/contents/articles/32711.exchange-online-how-to-enable-your-tenant-for-modern-authentication.aspx).|||UNTRANSLATED_CONTENT_END|||

## <a name="debugging-tools"></a>Ferramentas de depuração

Quando estiver desenvolvendo, é altamente recomendável que use uma ferramenta que possa interceptar e exibir as solicitações e respostas HTTP do serviço web do seu suplemento. Duas das mais populares são: 

- [Fiddler](http://www.telerik.com/fiddler): Gratuita ([Documentação](http://docs.telerik.com/fiddler/configure-fiddler/tasks/configurefiddler))
- [Charles](https://www.charlesproxy.com/): Gratuita por 30 dias. ([Documentação](https://www.charlesproxy.com/documentation/))

Ao desenvolver sua API de serviço, também pode tentar:

- [Postman](http://www.getpostman.com/postman): Gratuita ([Documentação](https://www.getpostman.com/docs/))

## <a name="causes-and-handling-of-errors-from-getaccesstokenasync"></a>Causas e tratamento dos erros do getAccessTokenAsync

Para acessar exemplos de manipulação dos erros descritos nesta seção, confira:
- [Home.js em Office-Add-in-ASPNET-SSO](https://github.com/OfficeDev/Office-Add-in-ASPNET-SSO/blob/master/Complete/Office-Add-in-ASPNET-SSO-WebAPI/Scripts/Home.js)
- [program.js em Office-Add-in-NodeJS-SSO](https://github.com/OfficeDev/Office-Add-in-NodeJS-SSO/blob/master/Completed/public/program.js)

> [!NOTE]
> Além das sugestões feitas nesta seção, um suplemento do Outlook tem uma outra maneira de responder a qualquer erro 13*nnn* . Para obter detalhes, consulte [cenário: implementar o logon único para seu serviço em um suplemento do Outlook](https://docs.microsoft.com/outlook/add-ins/implement-sso-in-outlook-add-in) e [o suplemento de amostra de AttachmentsDemo](https://github.com/OfficeDev/outlook-add-in-attachments-demo). 

### <a name="13000"></a>13000

A API [getAccessTokenAsync](https://docs.microsoft.com/office/dev/add-ins/develop/sso-in-office-add-ins#sso-api-reference) não é suportada pelo suplemento ou pela versão do Office. 

- A versão do Office não suporta SSO. A versão necessária é o Office 2016, versão 1710, compilação 8629.nnnn ou posterior (a versão de assinatura do Office 365, às vezes chamada de "Clique para Executar"). É possível que precise participar do programa Office Insider para obter essa versão. Para saber mais, confira a página [Seja um Office Insider](https://products.office.com/office-insider?tab=tab-1). 
- No manifesto do suplemento está faltando a seção [WebApplicationInfo](https://docs.microsoft.com/office/dev/add-ins/reference/manifest/webapplicationinfo?view=office-js) apropriada.

Seu suplemento deve responder a esse erro voltando para um sistema alternativo de autenticação do usuário. Para obter mais informações, consulte [requisitos e práticas recomendadas](https://docs.microsoft.com/office/dev/add-ins/develop/sso-in-office-add-ins#requirements-and-best-practices).

### <a name="13001"></a>13001

O usuário não tiver entrado no Office. Seu código deve se lembrar de `getAccessTokenAsync` método e passar a opção `forceAddAccount: true` no parâmetro [options](https://docs.microsoft.com/office/dev/add-ins/develop/sso-in-office-add-ins#sso-api-reference) . Mas não fazer isso mais de uma vez. O usuário pode ter decidido não entrar.

Este erro nunca é visto no Office Online. Se o cookie do usuário expirar, o Office Online retornará o erro 13006. 

### <a name="13002"></a>13002

O usuário cancelou a sessão ou consentiu; por exemplo, escolhendo **Cancelar** no diálogo de consentimento. 

- Se o suplemento fornece funções que não exigem que o usuário esteja conectado (ou que tenha concedido consentimento), seu código deve capturar esse erro e permitir que o suplemento permaneça em execução.
- Se o suplemento exige um usuário conectado que concedeu consentimento, seu código deverá solicitar ao usuário que repita a operação, mas não mais de uma vez. 

### <a name="13003"></a>13003

Tipo de usuário sem suporte. O usuário não está conectado no Office com uma conta do Office 365 ("trabalho ou escola") ou de Account válido da Microsoft. Isso pode acontecer se o Office for executado com uma conta de domínio local, por exemplo. Seu código seja pergunte ao usuário entrar no Office ou reverterá para um sistema alternativo de autenticação do usuário. Para obter mais informações, consulte [requisitos e práticas recomendadas](https://docs.microsoft.com/office/dev/add-ins/develop/sso-in-office-add-ins##requirements-and-best-practices).


### <a name="13004"></a>13004

Inválidos do recurso. O manifesto do suplemento ainda não foi configurado corretamente. Atualize o manifesto. Para obter mais informações, consulte [Validar e solucionar problemas com seu manifesto](../testing/troubleshoot-manifest.md). O máximo de comuns problema é que o elemento de **recurso** (no elemento **WebApplicationInfo** ) tem um domínio que não corresponde ao domínio do add-in. Embora a parte do protocolo do valor de recurso deve ser "api" não "https"; todas as outras partes do nome de domínio (incluindo porta, se houver alguma) devem ser o mesmo que para o suplemento.

### <a name="13005"></a>13005

Grant inválido. Geralmente, isso significa que o Office não foi pré-autorizados para web service do suplemento. Para obter mais informações, consulte [criar o aplicativo de serviço](sso-in-office-add-ins.md#create-the-service-application) e [registrar o suplemento com o ponto de extremidade do Azure AD v 2.0](create-sso-office-add-ins-aspnet.md#register-the-add-in-with-azure-ad-v20-endpoint) (ASP.NET) ou [registrar o suplemento com o ponto de extremidade do Azure AD v 2.0](create-sso-office-add-ins-nodejs.md#register-the-add-in-with-azure-ad-v20-endpoint) (nó JS). Isso também pode acontecer se o usuário não tiver concedido a suas permissões de aplicativo de serviço para seus `profile`.

### <a name="13006"></a>13006

Erro do cliente. Seu código deve sugerir para o usuário sair e reiniciar o Office ou reiniciar a sessão do Office Online.

### <a name="13007"></a>13007

O host do Office não conseguiu obter um token de acesso ao serviço Web do suplemento.

- Se este erro ocorre durante o desenvolvimento, certifique-se de que seu registro do suplemento e o suplemento manifesto especificam o `openid` e `profile` permissões. Para obter mais informações, consulte [registrar o suplemento com o ponto de extremidade do Azure AD v 2.0](create-sso-office-add-ins-aspnet.md#register-the-add-in-with-azure-ad-v20-endpoint) (ASP.NET) ou [registrar o suplemento com o ponto de extremidade do Azure AD v 2.0](create-sso-office-add-ins-nodejs.md#register-the-add-in-with-azure-ad-v20-endpoint) (nó JS) e [Configure o suplemento](create-sso-office-add-ins-aspnet.md#configure-the-add-in) (ASP.NET) ou [Configure o suplemento](create-sso-office-add-ins-nodejs.md#configure-the-add-in) (nó JS).
- Na produção, existem várias coisas que podem causar esse erro.
    - O usuário tem revogado consentimento, após anteriormente concedendo-lo. Seu código deve se lembrar de `getAccessTokenAsync` método com a opção `forceConsent: true`, mas não mais de uma vez.
    - O usuário é tem uma identidade de conta da Microsoft (MSA). Algumas situações em que pode causar um dos outros erros 13nnn com uma conta de trabalho ou escola, fará com que uma 13007 quando um MSA é usado. 

  Em todos esses casos, se você já tentou a opção `forceConsent` uma vez, então seu código poderia sugerir que o usuário tente novamente a operação mais tarde.

### <a name="13008"></a>13008

O usuário disparou uma operação que chama `getAccessTokenAsync` antes que uma chamada anterior de `getAccessTokenAsync` concluída. Seu código deve pedir ao usuário repita a operação após a operação anterior foi concluída.

### <a name="13009"></a>13009

O suplemento chamado o `getAccessTokenAsync` método com a opção `forceConsent: true`, mas o manifesto do add-in é implantado em um tipo de catálogo que não suporta forçando consentimento. Seu código deve se lembrar de `getAccessTokenAsync` método e passar a opção `forceConsent: false` no parâmetro [options](https://docs.microsoft.com/office/dev/add-ins/develop/sso-in-office-add-ins#sso-api-reference) . No entanto, a chamada de `getAccessTokenAsync` com `forceConsent: true` próprio pode ter sido uma resposta automática para uma chamada com falha de `getAccessTokenAsync` com `forceConsent: false`, portanto, seu código deve controlar se `getAccessTokenAsync` com `forceConsent: false` já foi chamado. Se ele tiver sido definido, seu código seja deve informar o usuário sair do Office e entrar novamente ou deve reverterá para um sistema alternativo de autenticação do usuário. Para obter mais informações, consulte [requisitos e práticas recomendadas](https://docs.microsoft.com/office/dev/add-ins/develop/sso-in-office-add-ins#requirements-and-best-practices).

> [!NOTE]
> Microsoft não necessariamente imporá essa restrição em todos os tipos de catálogos de suplemento. Caso contrário, esse erro nunca será visto.

### <a name="13010"></a>13010

O usuário está executando o add-in no Office Online e está usando o Internet Explorer ou borda. Domínio do Office 365 o usuário do e o domínio login.microsoftonline.com, estão em um zonas de segurança diferentes nas configurações do navegador. Se esse erro for retornado, o usuário será já vimos um erro explicando isso e vincular a uma página sobre como alterar a configuração de zona. Se o suplemento fornece funções que não exigem que o usuário estar conectado, seu código deve interceptar este erro e permitir que o suplemento permaneçam em execução.

### <a name="13012"></a>13012

O suplemento é executado em uma plataforma que não oferece suporte a `getAccessTokenAsync` API. Por exemplo, não é suportado no iPad. Consulte também [conjuntos de requisito de APIs de identidade](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/identity-api-requirement-sets).

### <a name="50001"></a>50001

Esse erro (que não é específica ao `getAccessTokenAsync`) pode indicar que o navegador tiver armazenado em cache uma cópia antiga dos arquivos do Office. js. Quando você estiver desenvolvendo, desmarque o cache do navegador. Também é possível que a versão do Office não é recente o suficiente para suportar o SSO. Consulte [pré-requisitos](create-sso-office-add-ins-aspnet.md#prerequisites).

Em um add-in de produção, o suplemento deve responder a esse erro voltando para um sistema alternativo de autenticação do usuário. Para obter mais informações, consulte [requisitos e práticas recomendadas](https://docs.microsoft.com/office/dev/add-ins/develop/sso-in-office-add-ins##requirements-and-best-practices).


## <a name="errors-on-the-server-side-from-azure-active-directory"></a>Erros no lado do servidor do Azure Active Directory

Para acessar exemplos de tratamento de erros descritos nesta seção, confira:
- [Office-Add-in-ASPNET-SSO](https://github.com/OfficeDev/Office-Add-in-ASPNET-SSO)
- [Office-Add-in-NodeJS-SSO](https://github.com/OfficeDev/Office-Add-in-NodeJS-SSO)


### <a name="conditional-access--multifactor-authentication-errors"></a>Erros no acesso condicional/autenticação multifatorial
 
Em determinadas configurações de identidade nos AAD e o Office 365, é possível que alguns recursos que estão acessíveis com o Microsoft Graph para exigir a autenticação multifator (MFA), mesmo quando não inquilino do Office 365 do usuário. Quando AAD recebe uma solicitação para obter um token para o recurso MFA protegidos, via fluxo em nome de, ele retorna ao serviço da web do seu suplemento uma mensagem JSON que contém um `claims` propriedade. A propriedade declarações tem informações sobre quais fatores adicionais de autenticação são necessários. 

Seu código do lado do servidor deve testar para esta mensagem e o valor de declarações para o seu código do lado do cliente de retransmissão. Você precisará dessas informações no cliente porque Office controla a autenticação para suplementos SSO. A mensagem para o cliente pode ser qualquer um erro (como `500 Server Error` ou `401 Unauthorized`) ou no corpo de uma resposta de sucesso (como `200 OK`). Em ambos os casos, o retorno de chamada (falha ou sucesso) do lado do cliente AJAX do seu código chamada para web da seu suplemento API deve ser testado para essa resposta. Se o valor de declarações tiver sido retransmitido, seu código deve lembrar `getAccessTokenAsync` e passe a opção `authChallenge: CLAIMS-STRING-HERE` no parâmetro [options](https://docs.microsoft.com/office/dev/add-ins/develop/sso-in-office-add-ins#sso-api-reference) . Quando AAD vê esta cadeia de caracteres, ele solicita ao usuário o factor(s) adicionais e, em seguida, retorna um novo token de acesso que serão aceitos no fluxo de em nome de.

### <a name="consent-missing-errors"></a>Erros de falta de consentimento

Se AAD não tiver nenhum registro que consentimento (para o recurso do Microsoft Graph) foi concedida para o suplemento pelo usuário (ou administrador de locatários), AAD enviará uma mensagem de erro para seu serviço da web. Seu código deve instruir o cliente (no corpo de um `403 Forbidden` resposta, por exemplo) para cancelamento `getAccessTokenAsync` com o `forceConsent: true` opção.

### <a name="invalid-or-missing-scope-permission-errors"></a>Erros de escopo (permissões) inválido ou ausente

- Seu código do lado do servidor deve enviar a resposta `403 Forbidden` ao cliente, que deve apresentar uma mensagem amigável ao usuário. Se possível, registre o erro no console ou registre-o em um log.
- Certifique-se de que seu suplemento [escopos](https://docs.microsoft.com/office/dev/add-ins/reference/manifest/scopes?view=office-js) seção manifesto especifica que todos necessárias permissões. E certifique-se de que seu registro do serviço de web do suplemento Especifica as mesmas permissões. Verifique se há erros de ortografia muito. Para obter mais informações, consulte [registrar o suplemento com o ponto de extremidade do Azure AD v 2.0](create-sso-office-add-ins-aspnet.md#register-the-add-in-with-azure-ad-v20-endpoint) (ASP.NET) ou [registrar o suplemento com o ponto de extremidade do Azure AD v 2.0](create-sso-office-add-ins-nodejs.md#register-the-add-in-with-azure-ad-v20-endpoint) (nó JS) e [Configure o suplemento](create-sso-office-add-ins-aspnet.md#configure-the-add-in) (ASP.NET) ou [Configure o suplemento](create-sso-office-add-ins-nodejs.md#configure-the-add-in) (nó JS).

### <a name="expired-or-invalid-token-errors-when-calling-microsoft-graph"></a>Erros de token expirados ou inválidos ao chamar o Microsoft Graph

Algumas bibliotecas de autenticação e autorização, inclusive MSAL, evitar erros de tokens expirados usando um token de atualização de cache sempre que necessário. Você também pode codificar seu próprio sistema de armazenamento em cache de token. Para obter um exemplo que faz isso, consulte [Office suplemento NodeJS SSO](https://github.com/OfficeDev/Office-Add-in-NodeJS-SSO), especialmente o arquivo [auth.ts](https://github.com/OfficeDev/Office-Add-in-NodeJS-SSO/blob/master/Completed/src/auth.ts).

Mas, se você recebe um token expirado ou um erro de token inválido, seu código deve dizer ao cliente (no corpo de uma resposta `401 Unauthorized`, por exemplo) para chamar novamente o `getAccessTokenAsync` e repetir a chamada para o ponto de extremidade da API Web do suplemento, que repetirá o fluxo on-behalf-of para obter um novo token para o Microsoft Graph. 

### <a name="invalid-token-error-when-calling-microsoft-graph"></a>Erro de token inválido ao chamar o Microsoft Graph

Trate esse erro da mesma forma que um erro de token expirado. Consulte a seção anterior.

### <a name="invalid-audience-error"></a>Erro de audiência inválida

Seu código do lado do servidor deve enviar uma resposta `403 Forbidden` ao cliente que apresente uma mensagem amigável ao usuário e, possivelmente, também registrar o erro no console ou gravá-lo em um registro.

Para obter mais informações sobre como adicionar suporte para validação de token por vários locatários, consulte a [Exemplo do Azure Multitenant](https://github.com/Azure-Samples/active-directory-dotnet-webapp-webapi-multitenant-openidconnect).
