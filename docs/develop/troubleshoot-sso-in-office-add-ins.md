---
title: Solucionar problemas de mensagens de erro no logon ?nico (SSO)
description: ''
ms.date: 12/08/2017
ms.openlocfilehash: 39099d746db3b5bea8a1ef629872006ba4ee087a
ms.sourcegitcommit: c72c35e8389c47a795afbac1b2bcf98c8e216d82
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/23/2018
---
# <a name="troubleshoot-error-messages-for-single-sign-on-sso-preview"></a>Solucionar problemas de mensagens de erro no logon ?nico (SSO) (visualiza??o)

Este artigo fornece algumas orienta??es sobre como solucionar problemas com o logon ?nico (SSO) nos suplementos do Office e como fazer com que seu suplemento habilitado para SSO lide de forma robusta com os erros ou condi??es especiais.

## <a name="debugging-tools"></a>Ferramentas de depura??o

Recomendamos fortemente que voc? use uma ferramenta que possa interceptar e exibir as solicita??es HTTP a partir de seu servi?o Web do suplemento, al?m de respostas para ele, quando voc? estiver desenvolvendo. Duas das ferramentas mais populares s?o: 

- [Fiddler](http://www.telerik.com/fiddler): gratuita ([Documenta??o](http://docs.telerik.com/fiddler/configure-fiddler/tasks/configurefiddler))
- [Charles](https://www.charlesproxy.com/): gratuita por 30 dias. ([Documenta??o](https://www.charlesproxy.com/documentation/))

Ao desenvolver sua API de servi?o, voc? tamb?m pode tentar:

- [Postman](http://www.getpostman.com/postman): Gratuita ([Documenta??o](https://www.getpostman.com/docs/))

## <a name="causes-and-handling-of-errors-from-getaccesstokenasync"></a>Causas e tratamento dos erros do getAccessTokenAsync

Para acessar exemplos de tratamento de erro descritos nesta se??o, confira:
- [Home.js em Office-Add-in-ASPNET-SSO](https://github.com/OfficeDev/Office-Add-in-ASPNET-SSO/blob/master/Complete/Office-Add-in-ASPNET-SSO-WebAPI/Scripts/Home.js)
- [program.js em Office-Add-in-NodeJS-SSO](https://github.com/OfficeDev/Office-Add-in-NodeJS-SSO/blob/master/Completed/public/program.js)

> [!NOTE]
> Al?m das sugest?es feitas nesta se??o, um suplemento do Outlook tem uma maneira adicional de responder a qualquer erro 13*nnn*. Para mais detalhes, consulte [Cen?rio: implementar o logon ?nico em seu servi?o em um suplemento do Outlook](https://docs.microsoft.com/en-us/outlook/add-ins/implement-sso-in-outlook-add-in) e [Suplemento de amostra AttachmentsDemo](https://github.com/OfficeDev/outlook-add-in-attachments-demo). 

### <a name="13000"></a>13000

A API [getAccessTokenAsync](https://dev.office.com/reference/add-ins/shared/office.context.auth.getAccessTokenAsync) n?o tem suporte do suplemento ou da vers?o do Office. 

- A vers?o do Office n?o d? suporte a SSO. Office 2016, vers?o 1710, build 8629.nnnn ou posterior (a vers?o de assinatura do Office 365, ?s vezes chamada de "Clique para Executar"). Talvez voc? precise ser um participante do programa Office Insider para obter essa vers?o. Para saber mais, confira a p?gina [Seja um Office Insider](https://products.office.com/en-us/office-insider?tab=tab-1). 
- O manifesto do suplemento est? sem a se??o [WebApplicationInfo](https://dev.office.com/reference/add-ins/manifest/webapplicationinfo) adequada.

### <a name="13001"></a>13001

O usu?rio n?o iniciou sess?o no Office. Seu c?digo deve chamar novamente o m?todo `getAccessTokenAsync` e passar a op??o `forceAddAccount: true` no par?metro [options](https://dev.office.com/reference/add-ins/shared/office.context.auth.getAccessTokenAsync#parameters). Mas n?o fa?a isso mais de uma vez. O usu?rio pode ter decidido n?o fazer login.

Este erro n?o ? visto no Office Online. Se os cookies do usu?rio expirarem, o Office Online retornar? o erro 13006. 

### <a name="13002"></a>13002

O usu?rio cancelou a sess?o ou consentiu; por exemplo, escolhendo **Cancelar** no di?logo de consentimento. 
- Se o seu suplemento fornece fun??es que n?o exigem que o usu?rio esteja conectado (ou que tenha concedido o consentimento), seu c?digo deve capturar esse erro e permitir que o suplemento permane?a em execu??o.
- Se o suplemento exigir um usu?rio conectado que concedeu consentimento, seu c?digo deve solicitar ao usu?rio que repita a opera??o, mas n?o mais do que uma vez. 

### <a name="13003"></a>13003

Tipo de Usu?rio n?o suportado. O usu?rio n?o iniciou sess?o no Office com uma conta Microsoft v?lida ou uma conta corporativa ou de estudante. Isso pode acontecer se o Office funcionar com uma conta de dom?nio no local, por exemplo. Seu c?digo deve solicitar ao usu?rio que fa?a login no Office.

### <a name="13004"></a>13004

Recurso inv?lido. O manifesto do suplemento n?o foi configurado corretamente. Atualize o manifesto. Para obter mais informa??es, consulte [Validar e solucionar problemas com seu manifesto](../testing/troubleshoot-manifest.md). O problema mais comum ? que o elemento **Recurso** (no elemento **WebApplicationInfo**) tem um dom?nio que n?o corresponde ao dom?nio do suplemento. Embora a parte do protocolo do valor do Recurso deva ser "api" e n?o "https"; todas as outras partes do nome de dom?nio (incluindo a porta, se houver) devem ser as mesmas do suplemento.

### <a name="13005"></a>13005

Concess?o inv?lida. Isso geralmente significa que o Office n?o foi pr?-autorizado para o servi?o Web do suplemento. Para obter mais informa??es, consulte [Criar o aplicativo de servi?o](sso-in-office-add-ins.md#create-the-service-application) e [Registrar o suplemento com o ponto de extremidade v2.0 do Azure AD](create-sso-office-add-ins-aspnet.md#register-the-add-in-with-azure-ad-v20-endpoint) (ASP.NET) ou [Registrar o suplemento com o ponto de extremidade v2.0 do Azure AD](create-sso-office-add-ins-nodejs.md#register-the-add-in-with-azure-ad-v20-endpoint) (N? JS). Isso tamb?m pode acontecer se o usu?rio n?o concedeu as permiss?es de aplicativo de servi?o para o seu `profile`.

### <a name="13006"></a>13006

Erro do cliente. Seu c?digo deve sugerir que o usu?rio saia e reinicie o Office ou reinicie a sess?o do Office Online.

### <a name="13007"></a>13007

O host do Office n?o conseguiu obter um token de acesso ao servi?o Web do suplemento.
- Se esse erro ocorrer durante o desenvolvimento, certifique-se de que o registro e o manifesto do suplemento especifiquem as permiss?es `openid` e `profile`. Para obter mais informa??es, consulte [Registrar o suplemento com o ponto de extremidade v2.0 do Azure AD](create-sso-office-add-ins-aspnet.md#register-the-add-in-with-azure-ad-v20-endpoint) (ASP.NET) ou [Registrar o suplemento com o ponto de extremidade v2.0 do Azure AD](create-sso-office-add-ins-nodejs.md#register-the-add-in-with-azure-ad-v20-endpoint) (N? JS), e [Configurar o suplemento](create-sso-office-add-ins-aspnet.md#configure-the-add-in) (ASP.NET) ou [Configurar o suplemento](create-sso-office-add-ins-nodejs.md#configure-the-add-in) (N? JS).
- Na produ??o, existem v?rias coisas que podem causar esse erro. Algumas s?o:
    - O usu?rio revogou o consentimento, ap?s conced?-lo anteriormente. Seu c?digo deve chamar novamente o `getAccessTokenAsync` m?todo com a op??o `forceConsent: true`, mas n?o mais que uma vez.
    - O usu?rio tem uma identidade de conta da Microsoft (MSA). Algumas situa??es que causariam um dos outros erros 13nnn com uma conta Work ou School causar?o um 13007 quando um MSA for usado. 

  Em todos esses casos, se voc? j? tiver tentado a op??o `forceConsent` uma vez, seu c?digo poder? sugerir que o usu?rio tente novamente a opera??o mais tarde.

### <a name="13008"></a>13008

O usu?rio desencadeou uma opera??o que chama o `getAccessTokenAsync` antes de uma chamada anterior do `getAccessTokenAsync` conclu?da. Seu c?digo deve solicitar ao usu?rio que repita a opera??o ap?s a conclus?o da opera??o anterior.

### <a name="13009"></a>13009

O suplemento chama o m?todo `getAccessTokenAsync` com a op??o `forceConsent: true`, mas o manifesto de suplemento foi implantado para um tipo de cat?logo n?o oferece suporte para for?ar o consentimento. Seu c?digo deve chamar novamente o m?todo `getAccessTokenAsync` e passar a op??o `forceConsent: false` no par?metro [options](https://dev.office.com/reference/add-ins/shared/office.context.auth.getAccessTokenAsync#parameters). No entanto, a chamada de `getAccessTokenAsync` com `forceConsent: true` pode ser uma resposta autom?tica para uma falha de chamada `getAccessTokenAsync` com `forceConsent: false`, assim o c?digo deve acompanhar se `getAccessTokenAsync` com `forceConsent: false` j? foi chamado. Em caso positivo, o c?digo deve informar para o usu?rio sair e entrar novamente no Office.

> [!NOTE]
> A Microsoft n?o imp?e necessariamente essa restri??o em quaisquer tipos de cat?logos de suplementos. Nesse caso, esse erro nunca ser? exibido.

### <a name="13010"></a>13010

O usu?rio est? executando o suplemento no Office Online e usando o Edge ou o Internet Explorer. O dom?nio do Office 365 do usu?rio e o dom?nio login.microsoftonline.com est?o em zonas de seguran?a diferentes nas configura??es do navegador. Se esse erro for retornado, o usu?rio j? ter? visto uma mensagem explicando o erro e vinculando a uma p?gina sobre como alterar a configura??o da zona. Se o seu suplemento fornece fun??es que n?o exigem que o usu?rio esteja conectado, o c?digo deve capturar esse erro e permitir que o suplemento permane?a em execu??o.

### <a name="50001"></a>50001

Este erro (que n?o ? espec?fico para `getAccessTokenAsync`) pode indicar que o navegador retirou uma c?pia antiga dos arquivos office.js. Limpe o cache do navegador. Outra possibilidade ? que a vers?o do Office n?o ? recente o suficiente para suportar o SSO. Consulte [Pr?-requisitos](create-sso-office-add-ins-aspnet.md#prerequisites).

## <a name="errors-on-the-server-side-from-azure-active-directory"></a>Erros do servidor do Active Directory do Azure

Para exemplos do tratamento de erro descritos nesta se??o, confira:
- [Office-Add-in-ASPNET-SSO](https://github.com/OfficeDev/Office-Add-in-ASPNET-SSO)
- [Office-Add-in-NodeJS-SSO](https://github.com/OfficeDev/Office-Add-in-NodeJS-SSO)


### <a name="conditional-access--multifactor-authentication-errors"></a>Erros no acesso condicional/autentica??o multifatorial
 
Em certas configura??es de identidade no AAD e no Office 365, ? poss?vel que alguns recursos que s?o acess?veis com o Microsoft Graph exijam autentica??o multifator (MFA), mesmo quando o locat?rio do Office 365 do usu?rio n?o exija. Quando o AAD recebe uma solicita??o de um token para o recurso protegido por MFA, atrav?s do fluxo Em Nome De, ele retorna ao servi?o Web do seu suplemento uma mensagem JSON que cont?m uma propriedade `claims`. A propriedade de reivindica??es tem informa??es sobre quais outros fatores de autentica??o s?o necess?rios. 

Seu c?digo do lado do servidor deve testar esta mensagem e transmitir o valor das reivindica??es ao seu c?digo do lado do cliente. Voc? precisa dessa informa??o no cliente porque o Office processa a autentica??o para os suplementos de SSO. A mensagem para o cliente pode ser um erro (como `500 Server Error` ou `401 Unauthorized`) ou estar no corpo de uma resposta de sucesso (como `200 OK`). Em ambos os casos, o retorno de chamada (falha ou sucesso) da chamada AJAX do lado do cliente do seu c?digo para a API da Web do seu suplemento deve testar essa resposta. Se o valor das solicita??es tiver sido retransmitido, seu c?digo deve chamar novamente o `getAccessTokenAsync` e passar a op??o `authChallenge: CLAIMS-STRING-HERE` no par?metro [op??es](https://dev.office.com/reference/add-ins/shared/office.context.auth.getAccessTokenAsync#parameters). Quando o AAD vir essa string, ele solicitar? ao usu?rio os fatores adicionais e retornar? um novo token de acesso que ser? aceito no fluxo Em Nome De.

### <a name="consent-missing-errors"></a>Erros de falta de consentimento

Se o AAD n?o tiver um registro de que o consentimento (para o recurso Microsoft Graph) foi concedido ao suplemento pelo usu?rio (ou administrador do locat?rio), o AAD enviar? uma mensagem de erro ao seu servi?o Web. Seu c?digo deve dizer ao cliente (no corpo de uma resposta `403 Forbidden`, por exemplo) para chamar novamente o `getAccessTokenAsync` com a op??o `forceConsent: true`.

### <a name="invalid-or-missing-scope-permission-errors"></a>Erros de escopos (permiss?es) inv?lidos ou ausentes

- Seu c?digo do lado do servidor deve enviar a resposta `403 Forbidden` ao cliente, que deve apresentar uma mensagem amig?vel ao usu?rio. Se poss?vel, registre o erro no console ou grave-o em um registro.
- Certifique-se de que a se??o de [Escopos](https://dev.office.com/reference/add-ins/manifest/scopes) do manifesto do seu suplemento especifique todas as permiss?es necess?rias. E certifique-se de que seu registro do servi?o Web do suplemento especifique as mesmas permiss?es. Verifique tamb?m os erros de ortografia. Para obter mais informa??es, consulte [Registrar o suplemento com o ponto de extremidade v2.0 do Azure AD](create-sso-office-add-ins-aspnet.md#register-the-add-in-with-azure-ad-v20-endpoint) (ASP.NET) ou [Registrar o suplemento com o ponto de extremidade v2.0 do Azure AD](create-sso-office-add-ins-nodejs.md#register-the-add-in-with-azure-ad-v20-endpoint) (N? JS), e [Configurar o suplemento](create-sso-office-add-ins-aspnet.md#configure-the-add-in) (ASP.NET) ou [Configurar o suplemento](create-sso-office-add-ins-nodejs.md#configure-the-add-in) (N? JS).

### <a name="expired-or-invalid-token-errors-when-calling-microsoft-graph"></a>Erros de token expirados ou inv?lidos ao chamar o Microsoft Graph

Algumas bibliotecas de autentica??o e autoriza??o, incluindo o MSAL, evitam erros de token expirados usando um token de atualiza??o em cache sempre que necess?rio. Voc? tamb?m pode codificar seu pr?prio sistema de cache de token. Para uma amostra que faz isso, consulte [Suplemento do Office NodeJS SSO](https://github.com/OfficeDev/Office-Add-in-NodeJS-SSO), especialmente o arquivo [auth.ts](https://github.com/OfficeDev/Office-Add-in-NodeJS-SSO/blob/master/Completed/src/auth.ts).

Mas se voc? receber um token expirado ou um erro de token inv?lido, seu c?digo deve dizer ao cliente (no corpo de uma resposta `401 Unauthorized`, por exemplo) para chamar novamente o `getAccessTokenAsync` e repetir a chamada para o ponto de extremidade de sua API da Web do suplemento, que repetir? o fluxo Em Nome De para obter um novo token para o Microsoft Graph. 

### <a name="invalid-token-error-when-calling-microsoft-graph"></a>Erro de token inv?lido ao chamar o Microsoft Graph

Lide com esse erro da mesma forma que um erro de token expirado. Consulte a se??o anterior.

### <a name="invalid-audience-error"></a>Erro de audi?ncia inv?lida

Seu c?digo do lado do servidor deve enviar uma resposta `403 Forbidden` ao cliente que deve apresentar uma mensagem amig?vel ao usu?rio e, possivelmente, tamb?m registrar o erro no console ou grav?-lo em um registro.

Para obter mais informa??es sobre a adi??o de suporte de v?rios locat?rios para valida??o de token, consulte a [Amostra de v?rios locat?rios do Azure](https://github.com/Azure-Samples/active-directory-dotnet-webapp-webapi-multitenant-openidconnect).
