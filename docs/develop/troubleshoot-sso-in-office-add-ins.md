---
title: Solucionar problemas de mensagens de erro no logon único (SSO)
description: ''
ms.date: 12/08/2017
ms.openlocfilehash: 270cc2c636f982d271f22fa93415515dbc63ad43
ms.sourcegitcommit: fdf7f4d686700edd6e6b04b2ea1bd43e59d4a03a
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 09/28/2018
ms.locfileid: "25348174"
---
# <a name="troubleshoot-error-messages-for-single-sign-on-sso-preview"></a>Solucionar problemas de mensagens de erro no logon único (SSO) (versão prévia)

Este artigo fornece algumas orientações sobre como solucionar problemas com o logon único (SSO) nos suplementos do Office e como fazer com que seu suplemento habilitado para SSO trate de forma robusta os erros ou condições especiais.

> [!NOTE]
> Atualmente a API de logon único tem suporte para Word, Excel e PowerPoint. Para obter mais informações sobre onde a API de logon único é suportada no momento, confira [conjuntos de requisitos IdentityAPI]https://docs.microsoft.com/javascript/office/requirement-sets/identity-api-requirement-sets).
> Para usar SSO, você precisa carregar a versão beta da Biblioteca JavaScript para Office de https://appsforoffice.microsoft.com/lib/beta/hosted/office.js na página HTML de inicialização do suplemento.
> Se você estiver trabalhando com um suplemento do Outlook, certifique-se de habilitar a Autenticação Moderna para a locação do Office 365. Confira mais informações sobre como fazer isso em [Exchange Online: como habilitar seu locatário para autenticação moderna](https://social.technet.microsoft.com/wiki/contents/articles/32711.exchange-online-how-to-enable-your-tenant-for-modern-authentication.aspx).

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
> Além das sugestões feitas nesta seção, um suplemento do Outlook tem uma maneira adicional de responder a qualquer erro 13*nnn*. Para detalhes, consulte [Cenário: implementar o logon único para seu serviço em um suplemento do Outlook](https://docs.microsoft.com/outlook/add-ins/implement-sso-in-outlook-add-in) e [Exemplo de Suplemento AttachmentsDemo](https://github.com/OfficeDev/outlook-add-in-attachments-demo). 

### <a name="13000"></a>13000

A API [getAccessTokenAsync](https://docs.microsoft.com/office/dev/add-ins/develop/sso-in-office-add-ins#sso-api-reference) não é suportada pelo suplemento ou pela versão do Office. 

- A versão do Office não suporta SSO. A versão necessária é o Office 2016, versão 1710, compilação 8629.nnnn ou posterior (a versão de assinatura do Office 365, às vezes chamada de "Clique para Executar"). É possível que precise participar do programa Office Insider para obter essa versão. Para saber mais, confira a página [Seja um Office Insider](https://products.office.com/office-insider?tab=tab-1). 
- No manifesto do suplemento está faltando a seção [WebApplicationInfo](https://docs.microsoft.com/javascript/office/manifest/webapplicationinfo?view=office-js) apropriada.

Seu suplemento deve responder a esse erro voltando para um sistema alternativo de autenticação de usuário. Para obter mais informações, consulte [Requisitos e práticas recomendadas](https://docs.microsoft.com/office/dev/add-ins/develop/sso-in-office-add-ins#requirements-and-best-practices).

### <a name="13001"></a>13001

O usuário não iniciou uma sessão no Office. Seu código deve chamar novamente o método `getAccessTokenAsync` e passar a opção `forceAddAccount: true` no parâmetro [options](https://docs.microsoft.com/office/dev/add-ins/develop/sso-in-office-add-ins#sso-api-reference). Porém, não faça isso mais de uma vez. O usuário pode ter decidido não fazer login.

Este erro nunca é visto no Office Online. Se o cookie do usuário expirar, o Office Online retornará o erro 13006. 

### <a name="13002"></a>13002

O usuário cancelou a sessão ou consentiu; por exemplo, escolhendo **Cancelar** no diálogo de consentimento. 

- Se o suplemento fornece funções que não exigem que o usuário esteja conectado (ou que tenha concedido consentimento), seu código deve capturar esse erro e permitir que o suplemento permaneça em execução.
- Se o suplemento exige um usuário conectado que concedeu consentimento, seu código deverá solicitar ao usuário que repita a operação, mas não mais de uma vez. 

### <a name="13003"></a>13003

Tipo de Usuário não suportado. O usuário não está conectado ao Office com uma conta válida da Microsoft ou do Office 365 ("Corporativa ou de Estudante"). Isso pode acontecer, por exemplo, se o Office for executado com uma conta de domínio local. Seu código deve pedir ao usuário para entrar no Office ou voltar a um sistema alternativo de autenticação de usuário. Para obter mais informações, consulte [Requisitos e práticas recomendadas](https://docs.microsoft.com/office/dev/add-ins/develop/sso-in-office-add-ins##requirements-and-best-practices).


### <a name="13004"></a>13004

Recurso inválido. O manifesto do suplemento não foi configurado corretamente. Atualizar o manifesto. Para obter mais informações, confira [Validar e solucionar problemas com seu manifesto](../testing/troubleshoot-manifest.md). O problema mais comum é que o elemento **Recurso** (no elemento **WebApplicationInfo**) tem um domínio que não corresponde ao domínio do suplemento. Apesar de que a parte do protocolo do valor do recurso deve ser "api" e não "https", todas as outras partes do nome de domínio (incluindo a porta, se houver) devem ser as mesmas do suplemento.

### <a name="13005"></a>13005

Concessão inválida. Isso geralmente significa que o Office não foi pré-autorizado para o serviço Web do suplemento. Para obter mais informações, consulte [Criar o aplicativo de serviço](sso-in-office-add-ins.md#create-the-service-application) e [Registrar o suplemento com o ponto de extremidade v2.0 do Azure AD](create-sso-office-add-ins-aspnet.md#register-the-add-in-with-azure-ad-v20-endpoint) (ASP.NET) ou [Registrar o suplemento com o ponto de extremidade v2.0 do Azure AD](create-sso-office-add-ins-nodejs.md#register-the-add-in-with-azure-ad-v20-endpoint) (Node JS). Isso também pode acontecer caso o usuário não tenha concedido as permissões de aplicativo de serviço para seu `profile`.

### <a name="13006"></a>13006

Erro do cliente. Seu código deve sugerir para o usuário sair e reiniciar o Office ou reiniciar a sessão do Office Online.

### <a name="13007"></a>13007

O host do Office não conseguiu obter um token de acesso ao serviço Web do suplemento.

- Se esse erro ocorrer durante o desenvolvimento, certifique-se de que o registro e o manifesto do suplemento especifiquem as permissões `openid` e `profile`. Para obter mais informações, consulte [Registrar o suplemento com o ponto de extremidade v2.0 do Azure AD](create-sso-office-add-ins-aspnet.md#register-the-add-in-with-azure-ad-v20-endpoint) (ASP.NET) ou [Registrar o suplemento com o ponto de extremidade v2.0 do Azure AD](create-sso-office-add-ins-nodejs.md#register-the-add-in-with-azure-ad-v20-endpoint) (Node JS), e [Configurar o suplemento](create-sso-office-add-ins-aspnet.md#configure-the-add-in) (ASP.NET) ou [Configurar o suplemento](create-sso-office-add-ins-nodejs.md#configure-the-add-in) (Node JS).
- Na produção, existem várias coisas que podem causar esse erro. Algumas são:
    - O usuário revogou o consentimento, após concedê-lo anteriormente. Seu código deve chamar novamente o `getAccessTokenAsync` método com a opção `forceConsent: true`, mas não mais que uma vez.
    - O usuário tem uma identidade de conta da Microsoft (MSA). Algumas situações que causariam um dos outros erros 13nnn com uma conta Work ou School causarão um 13007 quando um MSA for usado. 

  Em todos esses casos, se você já tentou a opção `forceConsent` uma vez, então seu código poderia sugerir que o usuário tente novamente a operação mais tarde.

### <a name="13008"></a>13008

O usuário acionou uma operação que chama o `getAccessTokenAsync` antes de uma chamada anterior concluída do `getAccessTokenAsync` . Seu código deve solicitar ao usuário que repita a operação após a conclusão da operação anterior.

### <a name="13009"></a>13009

O suplemento chamou o método `getAccessTokenAsync` com a opção `forceConsent: true`, mas o manifesto do suplemento é implantado para um tipo de catálogo que não oferece suporte para forçar o consentimento. Seu código deve chamar novamente o método `getAccessTokenAsync` e passar a opção `forceConsent: false` no parâmetro [options](https://docs.microsoft.com/office/dev/add-ins/develop/sso-in-office-add-ins#sso-api-reference). No entanto, a chamada de `getAccessTokenAsync` com `forceConsent: true` pode ser uma resposta automática para uma falha de chamada de `getAccessTokenAsync` com `forceConsent: false`, portanto, seu código deve monitorar se `getAccessTokenAsync` com `forceConsent: false` já foi chamado. Se foi, seu código deve instruir o usuário a sair do Office e entrar novamente ou deve retornar a um sistema alternativo de autenticação de usuário. Para obter mais informações, consulte [Requisitos e práticas recomendadas](https://docs.microsoft.com/office/dev/add-ins/develop/sso-in-office-add-ins#requirements-and-best-practices).

> [!NOTE]
> A Microsoft não impõe, necessariamente, essa restrição a nenhum tipo de catálogo de suplementos. Nesse caso, esse erro nunca será exibido.

### <a name="13010"></a>13010

O usuário está executando o suplemento no Office Online e usando o Edge ou o Internet Explorer. O domínio de usuário do Office 365 e o domínio login.microsoftonline.com estão em zonas de segurança diferentes nas configurações do navegador. Se retornar esse erro, o usuário verá uma mensagem explicando o erro com o vínculo para uma página sobre como alterar a configuração da zona. Se o seu suplemento fornece funções que não exigem que o usuário esteja conectado, o código deve capturar esse erro e permitir que o suplemento permaneça em execução.

### <a name="13012"></a>13012

O suplemento está sendo executado em uma plataforma que não oferece suporte à API `getAccessTokenAsync`. Por exemplo, não é suportado no iPad. Confira também [Conjuntos de requisitos de APIs de identidade](https://docs.microsoft.com/javascript/office/requirement-sets/identity-api-requirement-sets).

### <a name="50001"></a>50001

Esse erro (que não é específico de `getAccessTokenAsync`) pode indicar que o navegador armazenou em cache uma cópia antiga dos arquivos office.js. Quando você estiver desenvolvendo, desmarque o cache do navegador. Outra possibilidade é que a versão do Office esteja desatualizada e não suporte SSO. Consulte [Pré-requisitos](create-sso-office-add-ins-aspnet.md#prerequisites).

Em produção, o suplemento deve responder a esse erro retornando a um sistema alternativo de autenticação de usuário. Para obter mais informações, consulte [Requisitos e práticas recomendadas](https://docs.microsoft.com/office/dev/add-ins/develop/sso-in-office-add-ins##requirements-and-best-practices).


## <a name="errors-on-the-server-side-from-azure-active-directory"></a>Erros no servidor do Active Directory do Azure

Para acessar exemplos de tratamento de erros descritos nesta seção, confira:
- [Office-Add-in-ASPNET-SSO](https://github.com/OfficeDev/Office-Add-in-ASPNET-SSO)
- [Office-Add-in-NodeJS-SSO](https://github.com/OfficeDev/Office-Add-in-NodeJS-SSO)


### <a name="conditional-access--multifactor-authentication-errors"></a>Erros no acesso condicional/autenticação multifatorial
 
Em certas configurações de identidade no AAD e no Office 365, é possível que alguns recursos que podem ser acessados com o Microsoft Graph exijam autenticação multifatorial (MFA), mesmo que o locatário do usuário do Office 365 não a exija. Quando o AAD recebe uma solicitação de um token para o recurso protegido por MFA por meio do fluxo on-behalf-of, retorna ao serviço Web do suplemento uma mensagem JSON que contém uma propriedade `claims`. A propriedade de declarações possui informações sobre quais outros fatores de autenticação são necessários. 

Seu código servidor deve testar essa mensagem e transmitir o valor das declarações para o código do lado do cliente. Você precisa dessa informação no cliente porque o Office processa a autenticação para os suplementos de SSO. A mensagem para o cliente pode ser um erro (como `500 Server Error` ou `401 Unauthorized`) ou estar no corpo de uma resposta bem sucedida (como `200 OK`). Em ambos os casos, o retorno da chamada AJAX do seu código do lado do cliente (falha ou sucesso) para o suplemento da API Web deve testar essa resposta. Se o valor de claims for retransmitido, seu código deve chamar novamente o `getAccessTokenAsync` e passar a opção `authChallenge: CLAIMS-STRING-HERE` no parâmetro [options](https://docs.microsoft.com/office/dev/add-ins/develop/sso-in-office-add-ins#sso-api-reference). Quando o AAD vê essa sequência, solicita ao usuário os fatores adicionais e, em seguida, retorna um novo token de acesso que será aceito no fluxo on-behalf-of.

### <a name="consent-missing-errors"></a>Erros de falta de consentimento

Se o AAD não tiver um registro de que o consentimento (para o recurso do Microsoft Graph) foi concedido ao suplemento pelo usuário (ou administrador do locatário), o AAD enviará uma mensagem de erro ao seu serviço Web. Seu código deve dizer ao cliente (no corpo de uma resposta `403 Forbidden`, por exemplo) para chamar novamente o `getAccessTokenAsync` com a opção `forceConsent: true`.

### <a name="invalid-or-missing-scope-permission-errors"></a>Erros de escopo (permissões) inválido ou ausente

- Seu código do lado do servidor deve enviar a resposta `403 Forbidden` ao cliente, que deve apresentar uma mensagem amigável ao usuário. Se possível, registre o erro no console ou registre-o em um log.
- Certifique-se de que a seção do manifesto [Escopos](https://docs.microsoft.com/javascript/office/manifest/scopes?view=office-js)  do seu suplemento especifica todas as permissões necessárias. Certifique-se também de que o registro do serviço Web do suplemento especifica as mesmas permissões. Também verifique os erros de ortografia. Para obter mais informações, consulte [Registrar o suplemento com o ponto de extremidade v2.0 do Azure AD](create-sso-office-add-ins-aspnet.md#register-the-add-in-with-azure-ad-v20-endpoint) (ASP.NET) ou [Registrar o suplemento com o ponto de extremidade v2.0 do Azure AD](create-sso-office-add-ins-nodejs.md#register-the-add-in-with-azure-ad-v20-endpoint) (Node JS), e [Configurar o suplemento](create-sso-office-add-ins-aspnet.md#configure-the-add-in) (ASP.NET) ou [Configurar o suplemento](create-sso-office-add-ins-nodejs.md#configure-the-add-in) (Node JS).

### <a name="expired-or-invalid-token-errors-when-calling-microsoft-graph"></a>Erros de token expirados ou inválidos ao chamar o Microsoft Graph

Sempre que for necessário, algumas bibliotecas de autenticação e autorização, incluindo o MSAL, evitam erros de token expirados usando um token de atualização em cache. Você também pode codificar seu próprio sistema de cache de token. Para ver um exemplo disso, consulte [Suplemento do Office NodeJS SSO](https://github.com/OfficeDev/Office-Add-in-NodeJS-SSO), especialmente o arquivo [auth.ts](https://github.com/OfficeDev/Office-Add-in-NodeJS-SSO/blob/master/Completed/src/auth.ts).

Mas, se você recebe um token expirado ou um erro de token inválido, seu código deve dizer ao cliente (no corpo de uma resposta `401 Unauthorized`, por exemplo) para chamar novamente o `getAccessTokenAsync` e repetir a chamada para o ponto de extremidade da API Web do suplemento, que repetirá o fluxo on-behalf-of para obter um novo token para o Microsoft Graph. 

### <a name="invalid-token-error-when-calling-microsoft-graph"></a>Erro de token inválido ao chamar o Microsoft Graph

Trate esse erro da mesma forma que um erro de token expirado. Consulte a seção anterior.

### <a name="invalid-audience-error"></a>Erro de audiência inválida

Seu código do lado do servidor deve enviar uma resposta `403 Forbidden` ao cliente que apresente uma mensagem amigável ao usuário e, possivelmente, também registrar o erro no console ou gravá-lo em um registro.

Para obter mais informações sobre como adicionar suporte para validação de token por vários locatários, consulte a [Exemplo do Azure Multitenant](https://github.com/Azure-Samples/active-directory-dotnet-webapp-webapi-multitenant-openidconnect).
