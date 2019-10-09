---
title: Autorizar o Microsoft Graph com SSO
description: ''
ms.date: 08/09/2019
localization_priority: Priority
ms.openlocfilehash: 98b1219c0fe5459c497a27b915d31108545f14ae
ms.sourcegitcommit: 1dc1bb0befe06d19b587961da892434bd0512fb5
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 08/13/2019
ms.locfileid: "36302557"
---
# <a name="authorize-to-microsoft-graph-with-sso-preview"></a>Autorizar o Microsoft Graph com SSO (visualização)

Os usuários entram no Office (online, em dispositivos móveis e plataformas desktop) usando tanto a conta pessoal deles da Microsoft, como a conta corporativa ou de estudante (Office 365). A melhor maneira de um Suplemento do Office receber acesso autorizado ao [Microsoft Graph](https://developer.microsoft.com/graph/docs) é usar as credenciais de logon do Office do usuário. Isso permite a eles acessar seus dados do Microsoft Graph sem precisar entrar novamente. 

> [!NOTE]
> Atualmente a API de logon único tem suporte para Word, Excel e PowerPoint. Confira mais informações sobre os programas para os quais a API de logon único tem suporte no momento em [Conjuntos de requisitos da IdentityAPI](/office/dev/add-ins/reference/requirement-sets/identity-api-requirement-sets). Se você estiver trabalhando com um suplemento do Outlook, certifique-se de habilitar a Autenticação Moderna para a locação do Office 365. Confira mais informações sobre como fazer isso em [Exchange Online: como habilitar seu locatário para autenticação moderna](https://social.technet.microsoft.com/wiki/contents/articles/32711.exchange-online-how-to-enable-your-tenant-for-modern-authentication.aspx).

## <a name="add-in-architecture-for-sso-and-microsoft-graph"></a>Arquitetura de suplemento para SSO e Microsoft Graph

Além de hospedar as páginas e o JavaScript do aplicativo web, o suplemento também deve hospedar, ao mesmo tempo o [nome de domínio totalmente qualificado](/windows/desktop/DNS/f-gly#_dns_fully_qualified_domain_name_fqdn__gly), uma ou mais APIs web que obterá um token de acesso ao Microsoft Graph e fará solicitações a ele.

O manifesto do suplemento contém a marcação que especifica como ele está registrado no ponto de extremidade v2.0 do Azure Active Directory (Azure AD) e especifica todas as permissões para o Microsoft Graph que o suplemento precisa.

### <a name="how-it-works-at-runtime"></a>Como ele funciona em tempo de execução

O diagrama a seguir mostra como funciona o processo de entrar e obter acesso ao Microsoft Graph.

![Diagrama que mostra o processo de SSO](../images/sso-access-to-microsoft-graph.png)

1. No suplemento, o JavaScript chama uma nova API Office.js [getAccessTokenAsync](/office/dev/add-ins/develop/sso-in-office-add-ins#sso-api-reference). Isso informa ao aplicativo host do Office para obter um token de acesso para o suplemento. (De agora em diante, isso se chamará **token de acesso de inicialização** porque é substituído por um segundo token mais tarde durante o processo. Para ver um exemplo de um token de acesso de inicialização decodificado, confira [Token de acesso de exemplo](sso-in-office-add-ins.md#example-access-token).)
1. Se o usuário não estiver conectado, o aplicativo host do Office abrirá uma janela pop-up para o usuário entrar.
1. Se essa é a primeira vez que o usuário atual usa seu suplemento, será solicitado que ele dê o consentimento.
1. O aplicativo host do Office solicita o **token de acesso de inicialização** do ponto de extremidade v2.0 do Azure AD para o usuário atual.
1. O Azure AD envia o token de inicialização para o aplicativo host do Office.
1. O aplicativo host do Office envia o **token de acesso de inicialização** ao suplemento como parte do objeto de resultado retornado pela chamada de `getAccessTokenAsync`.
1. O JavaScript no suplemento faz uma solicitação HTTP a uma API Web que está hospedada no mesmo domínio totalmente qualificado que o suplemento e inclui o **token de acesso de inicialização** como prova de autorização.  
1. O código no lado do servidor valida o **token de acesso de inicialização** de entrada.
1. O código do lado do servidor usa o fluxo "on behalf of" (em nome de) (definido nos documentos [OAuth2 Token Exchange](https://tools.ietf.org/html/draft-ietf-oauth-token-exchange-02) e [Aplicativo para servidores ou daemon para um cenário com API web do Azure](/azure/active-directory/develop/active-directory-authentication-scenarios)) para obter um token de acesso para o Microsoft Graph em troca do token de acesso de inicialização.
1. O Azure AD retorna o token de acesso de inicialização para o Microsoft Graph (e um token de atualização, se o suplemento solicitar a permissão *offline_access*) para ele próprio.
1. O código do lado do servidor armazena em cache o token de acesso ao Microsoft Graph.
1. O código do lado do servidor faz solicitações ao Microsoft Graph e inclui o token de acesso ao Microsoft Graph.
1. O Microsoft Graph retorna os dados para o suplemento, que pode transmiti-los à interface do usuário do suplemento.
1. Quando o token de acesso ao Microsoft Graph expira, o código do lado do servidor pode usar o token de atualização para obter um novo token de acesso ao Microsoft Graph.

## <a name="develop-an-sso-add-in-that-accesses-microsoft-graph"></a>Desenvolver um suplemento SSO que acessa o Microsoft Graph

Você desenvolve um suplemento que acessa o Microsoft Graph como faria com qualquer outro suplemento que use SSO. Para obter uma descrição completa, confira [Habilitar o logon único para Suplementos do Office](/office/dev/add-ins/develop/sso-in-office-add-ins). A diferença é que é obrigatório que o suplemento tenha uma API Web do lado do servidor, e o token de acesso nesse artigo é chamado de "token de acesso de inicialização". 

Dependendo do seu idioma e da estrutura, podem estar disponíveis bibliotecas que simplificarão o código do lado do servidor que você precisa escrever. O código deve fazer o seguinte:

* Validar o token de acesso de inicialização que é recebido do manipulador de token que você criou anteriormente. Para saber mais, confira [Validar o token de acesso](sso-in-office-add-ins.md#validate-the-access-token). 
* Inicie o fluxo "on behalf of" com uma chamada para o ponto de extremidade v2.0 do Azure AD que inclui o token de acesso de inicialização, alguns metadados sobre o usuário e as credenciais do suplemento (sua ID e segredo).
* Armazene em cache o token de acesso retornado no Microsoft Graph. Para mais informações sobre esse fluxo, confira [Azure Active Directory v2.0 e fluxo "On-Behalf-Of" do OAuth 2.0](/azure/active-directory/develop/active-directory-v2-protocols-oauth-on-behalf-of).
* Crie um ou mais métodos de API Web que obtêm dados do Microsoft Graph passando o token de acesso em cache do Microsoft Graph.

> [!NOTE]
> Para exemplos de tokens de acesso decodificados do Microsoft Graph obtidos pelo fluxo "on behalf of", confira [Azure Active Directory v2.0 e fluxo "On-Behalf-Of" do OAuth 2.0](/azure/active-directory/develop/active-directory-v2-protocols-oauth-on-behalf-of).

Para obter exemplos detalhados passo a passo de cenários, confira:

* [Criar um Suplemento do Office com Node.js que usa logon único](create-sso-office-add-ins-nodejs.md)
* [Criar um Suplemento do Office com ASP.NET que usa logon único](create-sso-office-add-ins-aspnet.md)
* [Cenário: implementar o logon único no serviço em um suplemento do Outlook](/outlook/add-ins/implement-sso-in-outlook-add-in)
