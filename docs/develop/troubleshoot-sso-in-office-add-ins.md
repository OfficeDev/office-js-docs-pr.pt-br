---
title: Solucionar problemas de mensagens de erro no logon único (SSO)
description: Orientações sobre como solucionar problemas com logon único (SSO) em suplementos do Office e lidar com condições ou erros especiais.
ms.date: 04/13/2020
localization_priority: Normal
ms.openlocfilehash: da42b3b3d9b5cf1fede999a18bbe36c5532bd866
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 06/08/2020
ms.locfileid: "44609696"
---
# <a name="troubleshoot-error-messages-for-single-sign-on-sso-preview"></a>Solucionar problemas de mensagens de erro no logon único (SSO) (visualização)

Este artigo fornece algumas orientações sobre como solucionar problemas com o logon único (SSO) nos suplementos do Office e como fazer com que seu suplemento habilitado para SSO lide de forma robusta com os erros ou condições especiais.

> [!NOTE]
> Atualmente, a API de logon único tem suporte para Word, Excel, Outlook e PowerPoint. Confira mais informações sobre os programas para os quais a API de logon único tem suporte no momento em [Conjuntos de requisitos da IdentityAPI](../reference/requirement-sets/identity-api-requirement-sets.md).
> [!INCLUDE [Information about using preview APIs](../includes/using-excel-preview-apis.md)]
> Se você estiver trabalhando com um suplemento do Outlook, certifique-se de habilitar a Autenticação Moderna para a locação do Office 365. Confira mais informações sobre como fazer isso em [Exchange Online: como habilitar seu locatário para autenticação moderna](https://social.technet.microsoft.com/wiki/contents/articles/32711.exchange-online-how-to-enable-your-tenant-for-modern-authentication.aspx).

## <a name="debugging-tools"></a>Ferramentas de depuração

Recomendamos fortemente que você use uma ferramenta que possa interceptar e exibir as solicitações HTTP a partir de seu serviço Web do suplemento, além de respostas para ele, quando você estiver desenvolvendo. Duas das ferramentas mais populares são:

- [Fiddler](https://www.telerik.com/fiddler): gratuita ([Documentação](https://docs.telerik.com/fiddler/configure-fiddler/tasks/configurefiddler))
- [Charles](https://www.charlesproxy.com): Gratuita por 30 dias. ([Documentação](https://www.charlesproxy.com/documentation/))

## <a name="causes-and-handling-of-errors-from-getaccesstoken"></a>Causas e tratamento dos erros do getAccessToken

Para acessar exemplos de tratamento de erro descritos nesta seção, confira:
- [HomeES6.js em Office-Add-in-ASPNET-SSO](https://github.com/OfficeDev/Office-Add-in-ASPNET-SSO/blob/master/Complete/Office-Add-in-ASPNET-SSO-WebAPI/Scripts/HomeES6.js)
- [ssoAuthES6.js em Office-Add-in-NodeJS-SSO](https://github.com/OfficeDev/Office-Add-in-NodeJS-SSO/blob/master/Complete/public/javascripts/ssoAuthES6.js)

### <a name="13000"></a>13000

A API [getAccessToken](../develop/sso-in-office-add-ins.md#sso-api-reference) não é compatível pelo suplemento ou pela versão do Office.

- A versão do Office não é compatível com o SSO. A versão necessária é o Office 365 (a versão de assinatura do Office), em qualquer canal mensal.
- O manifesto do suplemento está sem a seção [WebApplicationInfo](../reference/manifest/webapplicationinfo.md) adequada.

O suplemento deverá responder a esse erro recorrendo a um sistema de autenticação de usuário alternativo. Para obter mais informações, confira [Requisitos e Melhores Práticas](../develop/sso-in-office-add-ins.md#requirements-and-best-practices).

### <a name="13001"></a>13001

O usuário não iniciou sessão no Office. Na maioria dos cenários, você deve evitar que esse erro apareça passando a opção `allowSignInPrompt: true` no parâmetro `AuthOptions`.

Mas pode haver exceções. Por exemplo, no caso de você desejar que o suplemento seja aberto com recursos que exijam um usuário conectado; mas *somente se* o usuário *já* estiver conectado ao Office. Se o usuário não estiver conectado, e você deseja que o suplemento seja aberto com um conjunto alternativo de recursos que não exijam que o usuário esteja. Nesse caso, essa é a lógica executada quando o suplemento inicia chamadas `getAccessToken` sem `allowSignInPrompt: true`. Use o erro 13001 como sinalizador para informar ao suplemento para apresentar o conjunto alternativo de recursos.

Outra opção é responder ao 13001 recorrendo a um sistema alternativo de autenticação de usuário. Isso conectará o usuário ao AAD, mas não o conectará ao Office.

Este erro nunca aparece no **Office na Web**. Se os cookies do usuário expirarem, o **Office na Web** retornará o erro 13006.

### <a name="13002"></a>13002

O usuário abortou a entrada ou o consentimento; por exemplo, escolhendo **Cancelar** no diálogo de consentimento.

- Se o seu suplemento fornece funções que não exigem que o usuário esteja conectado (ou que tenha concedido o consentimento), seu código deve capturar esse erro e permitir que o suplemento permaneça em execução.
- Se o suplemento exigir um usuário conectado que tenha dado consentimento, o código deverá exibir um botão de logon.

### <a name="13003"></a>13003

Tipo de Usuário não suportado. O usuário não iniciou sessão no Office com uma conta Microsoft ou do Office 365 válida (corporativa ou de estudante). Isso pode acontecer se o Office funcionar com uma conta de domínio no local, por exemplo. O código deve retornar a um sistema alternativo de autenticação de usuário. Para obter mais informações, confira [Requisitos e Melhores Práticas](../develop/sso-in-office-add-ins.md#requirements-and-best-practices).

### <a name="13004"></a>13004

Recurso inválido. (Esse erro só deve ser visto no desenvolvimento.) O manifesto do suplemento não foi configurado corretamente. Atualize o manifesto. Para saber mais, confira [Validar o manifesto de suplemento do Office](../testing/troubleshoot-manifest.md). O problema mais comum é que o elemento **Resource** (no elemento **WebApplicationInfo**) tem um domínio que não corresponde ao domínio do suplemento. Embora a parte do protocolo do valor Resource deva ser “api” e não “https”, todas as outras partes do nome de domínio (incluindo a porta, se houver) devem ser as mesmas para o suplemento.

### <a name="13005"></a>13005

Concessão inválida. Isso geralmente significa que o Office não foi pré-autorizado para o serviço Web do suplemento. Para obter mais informações, confira [Criar o aplicativo de serviço](sso-in-office-add-ins.md#create-the-service-application) e [Registrar o suplemento com o ponto de extremidade v2.0 do Azure AD](register-sso-add-in-aad-v2.md). Isso também pode acontecer se o usuário não concedeu as permissões de aplicativo de serviço para o seu `profile`, ou se tiver revogado um consentimento. O código deve retornar a um sistema alternativo de autenticação de usuário.

Outra causa possível durante o desenvolvimento, é que o suplemento esteja usando o Internet Explorer e você um certificado autoassinado. (Para determinar qual navegador está sendo usado em seu computador de desenvolvimento, confira [Navegadores usados pelos Suplementos do Offices](../concepts/browsers-used-by-office-web-add-ins.md).)

### <a name="13006"></a>13006

Erro do Cliente. Este erro somente aparece no **Office na Web**. Seu código deve sugerir que o usuário saia e reinicie a sessão do Office no navegador.

### <a name="13007"></a>13007

O host do Office não conseguiu obter um token de acesso ao serviço Web do suplemento.

- Se ocorrer este erro durante o desenvolvimento, certifique-se de que o registro e o manifesto do suplemento especifiquem as permissões de `profile` (e a permissão de `openid`, se estiver usando o MSAL.NET). Para mais informações, confira [Registrar o suplemento com o ponto de extremidade do Microsoft Azure AD v2.0](register-sso-add-in-aad-v2.md).
- Na produção, há várias coisas que podem causar esse erro. Algumas delas são:
    - O usuário tem uma identidade de Conta da Microsoft (MSA).
    - Algumas situações que causariam um dos outros erros 13xxx com uma conta Corporativa ou de Estudante, causariam um 13007 quando a MSA fosse usada.

  Em todos esses casos, o código deve retornar a um sistema alternativo de autenticação de usuário.

### <a name="13008"></a>13008

O usuário desencadeou uma operação que chama o `getAccessToken` antes de uma chamada anterior do `getAccessToken` concluída. Este erro somente aparece no **Office na Web**. O código deve solicitar ao usuário que repita a operação após a conclusão da operação anterior.

### <a name="13010"></a>13010

O usuário está executando o suplemento no Office, no Microsoft Edge ou no Internet Explorer. O domínio do Office 365 do usuário e o `login.microsoftonline.com` domínio estão em zonas de segurança diferentes nas configurações do navegador. Este erro somente aparece no **Office na Web**. Se esse erro for retornado, o usuário já terá visto uma mensagem explicando o erro e vinculando a uma página sobre como alterar a configuração da zona. Se o seu suplemento fornece funções que não exigem que o usuário esteja conectado, o código deve capturar esse erro e permitir que o suplemento permaneça em execução.

### <a name="13012"></a>13012

Há várias causas possíveis:

- O suplemento está em execução em uma plataforma não dá suporte à API `getAccessToken`. Por exemplo, ele não é suportado no iPad. Confira também [Conjuntos de requisitos da API de Identidade](../reference/requirement-sets/identity-api-requirement-sets.md).
- A opção `forMSGraphAccess` foi passada na chamada ao `getAccessToken` e o usuário obteve o suplemento no AppSource. Nesse cenário, o administrador do locatário não deu o consentimento ao suplemento para os escopos (permissões) do Microsoft Graph necessários. Uma nova chamada ao `getAccessToken` com o `allowConsentPrompt`, não resolverá o problema porque o Office poderá solicitar ao usuário o consentimento apenas para o escopo AAD do `profile`.

O código deve retornar a um sistema alternativo de autenticação de usuário.

No desenvolvimento, o suplemento é sideloaded no Outlook e a opção `forMSGraphAccess` foi passada na chamada ao `getAccessToken`.

### <a name="13013"></a>13013

O `getAccessToken` era chamado muitas vezes em um curto período de tempo, portanto, o Office limitou a chamada mais recente. Isso geralmente é causado por um loop infinito de chamadas para o método. Há situações em que se rechamar o método é aconselhável. No entanto, seu código deve usar uma variável de contador ou sinalizador para garantir que o método não seja rechamado repetidamente. Se o mesmo caminho de código de "repetição" estiver em execução novamente, o código deverá retornar a um sistema alternativo de autenticação do usuário. Para obter um exemplo de código, consulte como a `retryGetAccessToken` variável é usada no [HomeES6. js](https://github.com/OfficeDev/Office-Add-in-ASPNET-SSO/blob/master/Complete/Office-Add-in-ASPNET-SSO-WebAPI/Scripts/HomeES6.js) ou [ssoAuthES6. js](https://github.com/OfficeDev/Office-Add-in-NodeJS-SSO/blob/master/Complete/public/javascripts/ssoAuthES6.js).

### <a name="50001"></a>50001

Este erro (que não é específico de `getAccessToken`) pode indicar que o navegador colocou um cópia antiga dos arquivos office.js em cache. Quando você estiver desenvolvendo, limpe o cache do navegador. Também é possível que a versão do Office não esteja suficientemente recente para dar suporte à SSO. No Windows, a versão mínima é a 16.0.12215.20006. No Mac, é a 16.32.19102902.

Em um suplemento de produção, o suplemento deverá responder a esse erro recorrendo a um sistema de autenticação de usuário alternativo. Para obter mais informações, confira [Requisitos e Melhores Práticas](../develop/sso-in-office-add-ins.md#requirements-and-best-practices).

## <a name="errors-on-the-server-side-from-azure-active-directory"></a>Erros no lado do servidor do Azure Active Directory

Para exemplos do tratamento de erro descritos nesta seção, confira:
- [Office-Add-in-ASPNET-SSO](https://github.com/OfficeDev/Office-Add-in-ASPNET-SSO)
- [Office-Add-in-NodeJS-SSO](https://github.com/OfficeDev/Office-Add-in-NodeJS-SSO)

### <a name="conditional-access--multifactor-authentication-errors"></a>Erros no acesso condicional/autenticação multifatorial

Em certas configurações de identidade no AAD e no Office 365, é possível que alguns recursos que são acessíveis com o Microsoft Graph exijam autenticação multifator (MFA), mesmo quando o locatário do Office 365 do usuário não exija. Quando o AAD recebe uma solicitação de um token para o recurso protegido por MFA, através do fluxo Em Nome De, ele retorna ao serviço Web do seu suplemento uma mensagem JSON que contém uma propriedade `claims`. A propriedade de reivindicações tem informações sobre quais outros fatores de autenticação são necessários.

O código deve testar essa propriedade de `claims`. Dependendo da arquitetura do seu suplemento, você poderá testá-lo no lado do cliente ou testá-lo no lado do servidor e retransmiti-lo ao cliente. Você precisa dessa informação no cliente porque o Office processa a autenticação para os suplementos de SSO. Se você retransmiti-lo do lado do servidor, a mensagem para o cliente pode ser um erro (como `500 Server Error` ou `401 Unauthorized`) ou estar no corpo de uma resposta de sucesso (como `200 OK`). Em ambos os casos, o retorno de chamada (falha ou sucesso) da chamada AJAX do lado do cliente do seu código para a API da Web do seu suplemento deve testar essa resposta. 

Independentemente da arquitetura, se o valor de declarações tiver sido enviado do AAD, seu código deverá se lembrar `getAccessToken` e passar a opção `authChallenge: CLAIMS-STRING-HERE` no `options` parâmetro. Quando o AAD vir essa string, ele solicitará ao usuário os fatores adicionais e retornará um novo token de acesso que será aceito no fluxo Em Nome De.

### <a name="consent-missing-errors"></a>Erros de falta de consentimento

Se o AAD não tiver um registro de que o consentimento (para o recurso Microsoft Graph) foi concedido ao suplemento pelo usuário (ou administrador do locatário), o AAD enviará uma mensagem de erro ao seu serviço Web. Seu código deve dizer ao cliente (no corpo de uma resposta `403 Forbidden`, por exemplo).

Se o suplemento precisar de escopos do Microsoft Graph que só possam ser consentidos por um administrador, seu código deverá gerar um erro. Se os únicos escopos necessários puderem ser consentidos pelo usuário, o código deverá retornar a um sistema alternativo de autenticação de usuário.

### <a name="invalid-or-missing-scope-permission-errors"></a>Erros de escopos (permissões) inválidos ou ausentes

Esse tipo de erro só deve aparecer no desenvolvimento.

- Seu código do lado do servidor deve enviar a resposta `403 Forbidden` ao cliente, que deve registrar o erro no console ou gravá-lo em um log.
- Verifique se a seção de [Escopos](../reference/manifest/scopes.md) do manifesto do suplemento especifica todas as permissões necessárias. E certifique-se de que seu registro do serviço Web do suplemento especifique as mesmas permissões. Verifique também os erros de ortografia. Para mais informações, confira [Registrar o suplemento com o ponto de extremidade do Microsoft Azure AD v2.0](register-sso-add-in-aad-v2.md).

### <a name="invalid-audience-error-in-the-access-token-not-the-bootstrap-token"></a>Erro de audiência inválido no token de acesso (não no token de inicialização)

Seu código do lado do servidor deve enviar uma resposta `403 Forbidden` ao cliente que deve apresentar uma mensagem amigável ao usuário e, possivelmente, também registrar o erro no console ou gravá-lo em um registro.
