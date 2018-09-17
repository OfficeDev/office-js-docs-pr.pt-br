---
title: Solucionar problemas de mensagens de erro no logon único (SSO)
description: ''
ms.date: 12/08/2017
ms.openlocfilehash: 1dd36d99715937e12a9194baace3731ac331e658
ms.sourcegitcommit: 30435939ab8b8504c3dbfc62fd29ec6b0f1a7d22
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 09/12/2018
ms.locfileid: "23945436"
---
# <a name="troubleshoot-error-messages-for-single-sign-on-sso-preview"></a>Solucionar problemas de mensagens de erro no logon único (SSO) (visualização)

Este artigo fornece algumas orientações sobre como solucionar problemas com o logon único (SSO) nos suplementos do Office e como fazer com que seu suplemento habilitado para SSO lide de forma robusta com os erros ou condições especiais.

> [!NOTE]
> Atualmente a API de logon único tem suporte para Word, Excel e PowerPoint. Para obter mais informações sobre onde a API de logon único é suportada no momento, consulte [conjuntos de requisito IdentityAPI]https://docs.microsoft.com/javascript/office/requirement-sets/identity-api-requirement-sets).
> Para usar SSO, você precisa carregar a versão beta da Biblioteca JavaScript para Office de https://appsforoffice.microsoft.com/lib/beta/hosted/office.js na página HTML de inicialização do suplemento.
> Se você estiver trabalhando com um suplemento do Outlook, certifique-se de habilitar a Autenticação Moderna para a locação do Office 365. Confira mais informações sobre como fazer isso em [Exchange Online: como habilitar seu locatário para autenticação moderna](https://social.technet.microsoft.com/wiki/contents/articles/32711.exchange-online-how-to-enable-your-tenant-for-modern-authentication.aspx).

## <a name="debugging-tools"></a>Ferramentas de depuração

Recomendamos fortemente que você use uma ferramenta que possa interceptar e exibir as solicitações HTTP a partir de seu serviço Web do suplemento, além de respostas para ele, quando você estiver desenvolvendo. Duas das ferramentas mais populares são: 

- [Fiddler](http://www.telerik.com/fiddler): gratuita ([Documentação](http://docs.telerik.com/fiddler/configure-fiddler/tasks/configurefiddler))
- [Charles](https://www.charlesproxy.com/): gratuita por 30 dias. ([Documentação](https://www.charlesproxy.com/documentation/))

Ao desenvolver sua API de serviço, você também pode tentar:

- [Postman](http://www.getpostman.com/postman): Gratuita ([Documentação](https://www.getpostman.com/docs/))

## <a name="causes-and-handling-of-errors-from-getaccesstokenasync"></a>Causas e tratamento dos erros do getAccessTokenAsync

Para acessar exemplos de tratamento de erro descritos nesta seção, confira:
- [Home.js em Office-Add-in-ASPNET-SSO](https://github.com/OfficeDev/Office-Add-in-ASPNET-SSO/blob/master/Complete/Office-Add-in-ASPNET-SSO-WebAPI/Scripts/Home.js)
- [program.js em Office-Add-in-NodeJS-SSO](https://github.com/OfficeDev/Office-Add-in-NodeJS-SSO/blob/master/Completed/public/program.js)

> [!NOTE]
> Além das sugestões feitas nesta seção, um suplemento do Outlook tem uma maneira adicional de responder a qualquer erro 13*nnn*. Para detalhes, consulte [Cenário: implementar o logon único para seu serviço em um suplemento do Outlook](https://docs.microsoft.com/outlook/add-ins/implement-sso-in-outlook-add-in) e [Exemplo de Suplemento AttachmentsDemo](https://github.com/OfficeDev/outlook-add-in-attachments-demo). 

### <a name="13000"></a>13000

A API [getAccessTokenAsync](https://docs.microsoft.com/office/dev/add-ins/develop/sso-in-office-add-ins#sso-api-reference) não tem suporte do suplemento ou da versão do Office. 

- A versão do Office não dá suporte a SSO. Office 2016, versão 1710, build 8629.nnnn ou posterior (a versão de assinatura do Office 365, às vezes chamada de "Clique para Executar"). Talvez você precise ser um participante do programa Office Insider para obter essa versão. Para saber mais, confira a página [Seja um Office Insider](https://products.office.com/office-insider?tab=tab-1). 
- O manifesto do suplemento está sem a seção [WebApplicationInfo](https://docs.microsoft.com/javascript/office/manifest/webapplicationinfo?view=office-js) adequada.

### <a name="13001"></a>13001

O usuário não iniciou sessão no Office. Seu código deve chamar novamente o método `getAccessTokenAsync` e passar a opção `forceAddAccount: true` no parâmetro [options](https://docs.microsoft.com/office/dev/add-ins/develop/sso-in-office-add-ins#sso-api-reference). Mas não faça isso mais de uma vez. O usuário pode ter decidido não fazer login.

Este erro não é visto no Office Online. Se os cookies do usuário expirarem, o Office Online retornará o erro 13006. 

### <a name="13002"></a>13002

O usuário cancelou a sessão ou consentiu; por exemplo, escolhendo **Cancelar** no diálogo de consentimento. 
- Se o seu suplemento fornece funções que não exigem que o usuário esteja conectado (ou que tenha concedido o consentimento), seu código deve capturar esse erro e permitir que o suplemento permaneça em execução.
- Se o suplemento exigir um usuário conectado que concedeu consentimento, seu código deve solicitar ao usuário que repita a operação, mas não mais do que uma vez. 

### <a name="13003"></a>13003

Tipo de Usuário não suportado. O usuário não iniciou sessão no Office com uma conta Microsoft válida ou uma conta corporativa ou de estudante. Isso pode acontecer se o Office funcionar com uma conta de domínio no local, por exemplo. Seu código deve solicitar ao usuário que faça login no Office.

### <a name="13004"></a>13004

Recurso inválido. O manifesto do suplemento não foi configurado corretamente. Atualize o manifesto. Para obter mais informações, consulte [Validar e solucionar problemas com seu manifesto](../testing/troubleshoot-manifest.md). O problema mais comum é que o elemento **Recurso** (no elemento **WebApplicationInfo**) tem um domínio que não corresponde ao domínio do suplemento. Embora a parte do protocolo do valor do Recurso deva ser "api" e não "https"; todas as outras partes do nome de domínio (incluindo a porta, se houver) devem ser as mesmas do suplemento.

### <a name="13005"></a>13005

Concessão inválida. Isso geralmente significa que o Office não foi pré-autorizado para o serviço Web do suplemento. Para obter mais informações, consulte [Criar o aplicativo de serviço](sso-in-office-add-ins.md#create-the-service-application) e [Registrar o suplemento com o ponto de extremidade v2.0 do Azure AD](create-sso-office-add-ins-aspnet.md#register-the-add-in-with-azure-ad-v20-endpoint) (ASP.NET) ou [Registrar o suplemento com o ponto de extremidade v2.0 do Azure AD](create-sso-office-add-ins-nodejs.md#register-the-add-in-with-azure-ad-v20-endpoint) (Nó JS). Isso também pode acontecer caso o usuário não tenha concedido as permissões de aplicativo de serviço para seu `profile`.

### <a name="13006"></a>13006

Erro do cliente. Seu código deve sugerir que o usuário saia e reinicie o Office ou reinicie a sessão do Office Online.

### <a name="13007"></a>13007

O host do Office não conseguiu obter um token de acesso ao serviço Web do suplemento.
- Se esse erro ocorrer durante o desenvolvimento, certifique-se de que o registro e o manifesto do suplemento especifiquem as permissões `openid` e `profile`. Para obter mais informações, consulte [Registrar o suplemento com o ponto de extremidade v2.0 do Azure AD](create-sso-office-add-ins-aspnet.md#register-the-add-in-with-azure-ad-v20-endpoint) (ASP.NET) ou [Registrar o suplemento com o ponto de extremidade v2.0 do Azure AD](create-sso-office-add-ins-nodejs.md#register-the-add-in-with-azure-ad-v20-endpoint) (Nó JS), e [Configurar o suplemento](create-sso-office-add-ins-aspnet.md#configure-the-add-in) (ASP.NET) ou [Configurar o suplemento](create-sso-office-add-ins-nodejs.md#configure-the-add-in) (Nó JS).
- Na produção, existem várias coisas que podem causar esse erro. Algumas são:
    - O usuário revogou o consentimento, após concedê-lo anteriormente. Seu código deve chamar novamente o `getAccessTokenAsync` método com a opção `forceConsent: true`, mas não mais que uma vez.
    - O usuário tem uma identidade de conta da Microsoft (MSA). Algumas situações que causariam um dos outros erros 13nnn com uma conta Work ou School causarão um 13007 quando um MSA for usado. 

  Em todos esses casos, se você já tiver tentado a opção `forceConsent` uma vez, seu código poderá sugerir que o usuário tente novamente a operação mais tarde.

### <a name="13008"></a>13008

O usuário desencadeou uma operação que chama o `getAccessTokenAsync` antes de uma chamada anterior do `getAccessTokenAsync` concluída. Seu código deve solicitar ao usuário que repita a operação após a conclusão da operação anterior.

### <a name="13009"></a>13009

O suplemento chama o método `getAccessTokenAsync` com a opção `forceConsent: true`, mas o manifesto de suplemento foi implantado para um tipo de catálogo não oferece suporte para forçar o consentimento. Seu código deve chamar novamente o método `getAccessTokenAsync` e passar a opção `forceConsent: false` no parâmetro [options](https://docs.microsoft.com/office/dev/add-ins/develop/sso-in-office-add-ins#sso-api-reference). No entanto, a chamada de `getAccessTokenAsync` com `forceConsent: true` pode ser uma resposta automática para uma falha de chamada `getAccessTokenAsync` com `forceConsent: false`, assim o código deve acompanhar se `getAccessTokenAsync` com `forceConsent: false` já foi chamado. Em caso positivo, o código deve informar para o usuário sair e entrar novamente no Office.

> [!NOTE]
> A Microsoft não impõe, necessariamente, essa restrição a nenhum tipo de catálogo de suplementos. Nesse caso, esse erro nunca será exibido.

### <a name="13010"></a>13010

O usuário está executando o suplemento no Office Online e usando o Edge ou o Internet Explorer. O domínio do Office 365 do usuário e o domínio login.microsoftonline.com estão em zonas de segurança diferentes nas configurações do navegador. Se esse erro for retornado, o usuário já terá visto uma mensagem explicando o erro e vinculando a uma página sobre como alterar a configuração da zona. Se o seu suplemento fornece funções que não exigem que o usuário esteja conectado, o código deve capturar esse erro e permitir que o suplemento permaneça em execução.

### <a name="50001"></a>50001

Este erro (que não é específico para `getAccessTokenAsync`) pode indicar que o navegador retirou uma cópia antiga dos arquivos office.js. Limpe o cache do navegador. Outra possibilidade é que a versão do Office não é recente o suficiente para suportar o SSO. Consulte [Pré-requisitos](create-sso-office-add-ins-aspnet.md#prerequisites).

## <a name="errors-on-the-server-side-from-azure-active-directory"></a>Erros no lado do servidor do Azure Active Directory

Para exemplos do tratamento de erro descritos nesta seção, confira:
- [Office-Add-in-ASPNET-SSO](https://github.com/OfficeDev/Office-Add-in-ASPNET-SSO)
- [Office-Add-in-NodeJS-SSO](https://github.com/OfficeDev/Office-Add-in-NodeJS-SSO)


### <a name="conditional-access--multifactor-authentication-errors"></a>Erros no acesso condicional/autenticação multifatorial
 
Em certas configurações de identidade no AAD e no Office 365, é possível que alguns recursos que são acessíveis com o Microsoft Graph exijam autenticação multifator (MFA), mesmo quando o locatário do Office 365 do usuário não exija. Quando o AAD recebe uma solicitação de um token para o recurso protegido por MFA, através do fluxo Em Nome De, ele retorna ao serviço Web do seu suplemento uma mensagem JSON que contém uma propriedade `claims`. A propriedade de reivindicações tem informações sobre quais outros fatores de autenticação são necessários. 

Seu código do lado do servidor deve testar esta mensagem e transmitir o valor das reivindicações ao seu código do lado do cliente. Você precisa dessa informação no cliente porque o Office processa a autenticação para os suplementos de SSO. A mensagem para o cliente pode ser um erro (como `500 Server Error` ou `401 Unauthorized`) ou estar no corpo de uma resposta de sucesso (como `200 OK`). Em ambos os casos, o retorno de chamada (falha ou sucesso) da chamada AJAX do lado do cliente do seu código para a API da Web do seu suplemento deve testar essa resposta. Se o valor das solicitações tiver sido retransmitido, seu código deve chamar novamente o `getAccessTokenAsync` e passar a opção `authChallenge: CLAIMS-STRING-HERE` no parâmetro [opções](https://docs.microsoft.com/office/dev/add-ins/develop/sso-in-office-add-ins#sso-api-reference). Quando o AAD vir essa string, ele solicitará ao usuário os fatores adicionais e retornará um novo token de acesso que será aceito no fluxo Em Nome De.

### <a name="consent-missing-errors"></a>Erros de falta de consentimento

Se o AAD não tiver um registro de que o consentimento (para o recurso Microsoft Graph) foi concedido ao suplemento pelo usuário (ou administrador do locatário), o AAD enviará uma mensagem de erro ao seu serviço Web. Seu código deve dizer ao cliente (no corpo de uma resposta `403 Forbidden`, por exemplo) para chamar novamente o `getAccessTokenAsync` com a opção `forceConsent: true`.

### <a name="invalid-or-missing-scope-permission-errors"></a>Erros de escopos (permissões) inválidos ou ausentes

- Seu código do lado do servidor deve enviar a resposta `403 Forbidden` ao cliente, que deve apresentar uma mensagem amigável ao usuário. Se possível, registre o erro no console ou grave-o em um registro.
- Certifique-se de que a seção de [Escopos](https://docs.microsoft.com/javascript/office/manifest/scopes?view=office-js) do manifesto do seu suplemento especifique todas as permissões necessárias. E certifique-se de que seu registro do serviço Web do suplemento especifique as mesmas permissões. Verifique também os erros de ortografia. Para obter mais informações, consulte [Registrar o suplemento com o ponto de extremidade v2.0 do Azure AD](create-sso-office-add-ins-aspnet.md#register-the-add-in-with-azure-ad-v20-endpoint) (ASP.NET) ou [Registrar o suplemento com o ponto de extremidade v2.0 do Azure AD](create-sso-office-add-ins-nodejs.md#register-the-add-in-with-azure-ad-v20-endpoint) (Nó JS), e [Configurar o suplemento](create-sso-office-add-ins-aspnet.md#configure-the-add-in) (ASP.NET) ou [Configurar o suplemento](create-sso-office-add-ins-nodejs.md#configure-the-add-in) (Nó JS).

### <a name="expired-or-invalid-token-errors-when-calling-microsoft-graph"></a>Erros de token expirados ou inválidos ao chamar o Microsoft Graph

Algumas bibliotecas de autenticação e autorização, incluindo o MSAL, evitam erros de token expirados usando um token de atualização em cache sempre que necessário. Você também pode codificar seu próprio sistema de cache de token. Para uma amostra que faz isso, consulte [Suplemento do Office NodeJS SSO](https://github.com/OfficeDev/Office-Add-in-NodeJS-SSO), especialmente o arquivo [auth.ts](https://github.com/OfficeDev/Office-Add-in-NodeJS-SSO/blob/master/Completed/src/auth.ts).

Mas se você receber um token expirado ou um erro de token inválido, seu código deve dizer ao cliente (no corpo de uma resposta `401 Unauthorized`, por exemplo) para chamar novamente o `getAccessTokenAsync` e repetir a chamada para o ponto de extremidade de sua API da Web do suplemento, que repetirá o fluxo Em Nome De para obter um novo token para o Microsoft Graph. 

### <a name="invalid-token-error-when-calling-microsoft-graph"></a>Erro de token inválido ao chamar o Microsoft Graph

Lide com esse erro da mesma forma que um erro de token expirado. Consulte a seção anterior.

### <a name="invalid-audience-error"></a>Erro de audiência inválida

Seu código do lado do servidor deve enviar uma resposta `403 Forbidden` ao cliente que deve apresentar uma mensagem amigável ao usuário e, possivelmente, também registrar o erro no console ou gravá-lo em um registro.

Para obter mais informações sobre a adição de suporte de vários locatários para validação de token, consulte a [Amostra de vários locatários do Azure](https://github.com/Azure-Samples/active-directory-dotnet-webapp-webapi-multitenant-openidconnect).
