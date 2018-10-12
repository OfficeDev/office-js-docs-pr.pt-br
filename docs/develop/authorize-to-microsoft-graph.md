---
title: Autorizar para o Microsoft Graph no seu Suplemento do Office
description: ''
ms.date: 04/10/2018
ms.openlocfilehash: 6d0b6f2002b71c4680b72d2f40492fff1abf15e2
ms.sourcegitcommit: c53f05bbd4abdfe1ee2e42fdd4f82b318b363ad7
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 10/12/2018
ms.locfileid: "25505853"
---
# <a name="authorize-to-microsoft-graph-in-your-office-add-in-preview"></a>Autorizar para o Microsoft Graph no seu Suplemento do Office (visualização)

Os usuários entram no Office (plataformas online, móveis e desktop) usando sua conta pessoal da Microsoft ou sua conta corporativa ou de estudante (Office 365). A melhor maneira de um Suplemento do Office ter acesso autorizado ao [Microsoft Graph](https://developer.microsoft.com/graph/docs) é usar as credenciais de entrada do usuário no Office. Isso permite acessar os dados do Microsoft Graph sem precisar entrar uma segunda vez. 

> [!NOTE]
> A API de logon único é atualmente compatível com as versões prévias do Word, Excel, Outlook e PowerPoint. Para mais informações sobre a compatibilidade da API de logon único, veja [Conjuntos de requisitos da API de identidade](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/identity-api-requirement-sets?view=office-js). Se você estiver trabalhando em um suplemento do Outlook, não esqueça de habilitar a Autenticação Moderna para o locatário do Office 365. Para saber como fazer isso, veja [Exchange Online: Como habilitar o seu locatário para a Autenticação Moderna](https://social.technet.microsoft.com/wiki/contents/articles/32711.exchange-online-how-to-enable-your-tenant-for-modern-authentication.aspx).

## <a name="add-in-architecture-for-sso-and-microsoft-graph"></a>Arquitetura do suplemento para SSO e Microsoft Graph

Além de hospedar as páginas e o JavaScript do aplicativo da Web, o suplemento também deve hospedar, no mesmo [nome de domínio totalmente qualificado](https://docs.microsoft.com/windows/desktop/DNS/f-gly#_dns_fully_qualified_domain_name_fqdn__gly), uma ou mais APIs web que obterão um token de acesso para o Microsoft Graph e farão solicitações para ele.

O manifesto do suplemento contém a marcação que especifica como ele está registrado no ponto de extremidade v2.0 do Active Directory do Azure (AD do Azure) e especifica todas as permissões para o Microsoft Graph que o suplemento precisa.

### <a name="how-it-works-at-runtime"></a>Como funciona em tempo de execução

O diagrama a seguir mostra como o processo de entrada e acesso ao Microsoft Graph funciona.

![Diagrama que mostra o processo de SSO](../images/sso-access-to-microsoft-graph.png)

1. No suplemento, o JavaScript chama uma nova API Office.js [getAccessTokenAsync](https://docs.microsoft.com/office/dev/add-ins/develop/sso-in-office-add-ins#sso-api-reference). Isso informa ao aplicativo host do Office para obter um token de acesso para o suplemento. (Daqui por diante, ele será chamado de **token de acesso de inicialização** porque será substituído por um segundo token no processo posteriormente. Para ver um exemplo de um token de acesso de inicialização decodificado, veja [Exemplo de token de acesso](sso-in-office-add-ins.md#example-access-token).)
1. Se o usuário não estiver conectado, o aplicativo host do Office abrirá uma janela pop-up para o usuário entrar.
1. Se essa for a primeira vez que o usuário atual usa o suplemento, ele será solicitado a consentir.
1. O aplicativo host do Office solicita o **token de acesso de inicialização** do ponto de extremidade do Azure AD v2.0 para o usuário atual.
1. O Azure AD envia o token de inicialização ao aplicativo host do Office.
1. O aplicativo host do Office envia o **token de acesso de inicialização** ao suplemento como parte do objeto de resultado que retornou pela chamada de `getAccessTokenAsync`.
1. O JavaScript no suplemento faz uma solicitação HTTP a uma API da Web que está hospedada no mesmo domínio totalmente qualificado que o suplemento e inclui o **token de acesso de inicialização** como prova de autorização.  
1. O código do servidor valida o **token de acesso de inicialização** de entrada.
1. O código do servidor usa o fluxo “em nome de” (definido em [Troca de Token do OAuth2](https://tools.ietf.org/html/draft-ietf-oauth-token-exchange-02) e o [daemon ou aplicativo para servidores para o cenário da API da Web do Azure](https://docs.microsoft.com/azure/active-directory/develop/active-directory-authentication-scenarios#daemon-or-server-application-to-web-api)) para obter um token de acesso para o Microsoft Graph em troca do token de acesso de inicialização.
1. O Azure AD retorna o token de acesso ao Microsoft Graph (e um token de atualização, se o suplemento solicitar a permissão *offline_access*) ao suplemento.
1. O código do servidor armazena em cache o token de acesso ao Microsoft Graph.
1. O código do lado do servidor faz solicitações ao Microsoft Graph e inclui o token de acesso ao Microsoft Graph.
1. O Microsoft Graph retorna os dados para o suplemento, que pode transmiti-los à interface do usuário do suplemento.
1. Quando o token de acesso ao Microsoft Graph expira, o código do servidor pode usar seu token de atualização para obter um novo token de acesso ao Microsoft Graph.

## <a name="develop-an-sso-add-in-that-accesses-microsoft-graph"></a>Desenvolver um suplemento SSO que acesse o Microsoft Graph

Desenvolver um suplemento que acessa o Microsoft Graph é exatamente como desenvolver qualquer outro suplemento que usa SSO. Para obter uma descrição detalhada, consulte [Habilitar o logon único para suplementos do Office](https://docs.microsoft.com/office/dev/add-ins/develop/sso-in-office-add-ins). A diferença é que é obrigatório que o suplemento tenha uma API Web do lado do servidor, e o que é chamado de token de acesso neste artigo é chamado de "token de acesso de inicialização". 

Dependendo da sua linguagem e estrutura, pode haver bibliotecas disponíveis para simplificar o código do lado do servidor de que você precisa. Seu código deve fazer o seguinte:

* Validar o token de acesso de inicialização do suplemento que é recebido do manipulador de tokens que você criou anteriormente. Para obter mais informações, consulte [Validar o token de acesso](sso-in-office-add-ins.md#validate-the-access-token). 
* Iniciar o fluxo "em nome de" com uma chamada para o ponto de extremidade v2.0 do AD do Azure que inclui o token de acesso de inicialização, alguns metadados sobre o usuário e as credenciais do suplemento (ID e segredo).
* Armazena em cache o token de acesso retornado para o Microsoft Graph. Para obter mais informações sobre esse fluxo consulte [v 2.0 do Active Directory do Azure e o fluxo "em nome de" do OAuth 2.0](https://docs.microsoft.com/azure/active-directory/develop/active-directory-v2-protocols-oauth-on-behalf-of).
* Criar um ou mais métodos de API da Web que obtenham dados do Microsoft Graph passando o token de acesso em cache para o Microsoft Graph.

> [!NOTE]
> Para ver exemplos de tokens de acesso decodificados para o Microsoft Graph e obtidos pelo fluxo "em nome de", veja [Active Directory do Azure v2.0v e "fluxo em nome de" do OAuth 2.0](https://docs.microsoft.com/azure/active-directory/develop/active-directory-v2-protocols-oauth-on-behalf-of).

Confira exemplos de passo a passo e cenários detalhados em:

* [Criar um Suplemento do Office com Node.js que usa logon único](create-sso-office-add-ins-nodejs.md)
* [Criar um Suplemento do Office com ASP.NET que usa logon único](create-sso-office-add-ins-aspnet.md)
* [Cenário: implementar o logon único no serviço em um suplemento do Outlook](https://docs.microsoft.com/outlook/add-ins/implement-sso-in-outlook-add-in)



