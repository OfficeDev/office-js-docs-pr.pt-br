---
title: Autorizar para o Microsoft Graph no seu Suplemento do Office
description: ''
ms.date: 04/10/2018
ms.openlocfilehash: 495aa5554550d10711c418339d412e3a312d02fb
ms.sourcegitcommit: c72c35e8389c47a795afbac1b2bcf98c8e216d82
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/23/2018
---
# <a name="authorize-to-microsoft-graph-in-your-office-add-in-preview"></a>Autorizar para o Microsoft Graph no seu Suplemento do Office (visualiza??o)

Os usu?rios entram no Office (online, em dispositivos m?veis e plataformas desktop) usando tanto a conta pessoal deles da Microsoft, como a conta corporativa ou de estudante (Office 365). A melhor maneira de um Suplemento do Office ter acesso autorizado ao [Microsoft Graph](https://developer.microsoft.com/graph/docs) ? usar as credenciais de entrada do usu?rio no Office. Isso permite acessar os dados do Microsoft Graph sem precisar entrar uma segunda vez. 

> [!NOTE]
> Atualmente, a API de logon ?nico tem suporte para Word, Excel e PowerPoint. Confira mais informa??es sobre os programas para os quais a API de logon ?nico tem suporte no momento em [Conjuntos de requisitos da IdentityAPI](https://dev.office.com/reference/add-ins/requirement-sets/identity-api-requirement-sets).
> Se voc? estiver trabalhando com um suplemento do Outlook, certifique-se de habilitar a Autentica??o Moderna para a loca??o do Office 365. Confira mais informa??es sobre como fazer isso em [Exchange Online: como habilitar seu locat?rio para autentica??o moderna](https://social.technet.microsoft.com/wiki/contents/articles/32711.exchange-online-how-to-enable-your-tenant-for-modern-authentication.aspx).

## <a name="add-in-architecture-for-sso-and-microsoft-graph"></a>Arquitetura do suplemento para SSO e Microsoft Graph

Al?m de hospedar as p?ginas e o JavaScript do aplicativo web, o suplemento tamb?m deve hospedar, ao mesmo tempo o [nome de dom?nio totalmente qualificado](https://msdn.microsoft.com/en-us/library/windows/desktop/ms682135.aspx#_dns_fully_qualified_domain_name_fqdn__gly), uma ou mais APIs web que obter? um token de acesso ao Microsoft Graph e far? solicita??es a ele.

O manifesto do suplemento cont?m a marca??o que especifica como ele est? registrado no ponto de extremidade v2.0 do Azure Active Directory (Azure AD) e especifica todas as permiss?es para o Microsoft Graph que o suplemento precisa.

### <a name="how-it-works-at-runtime"></a>Como ele funciona em tempo de execu??o

O diagrama a seguir mostra como o processo de entrada e acesso ao Microsoft Graph funciona.

![Diagrama que mostra o processo de SSO](../images/sso-access-to-microsoft-graph.png)

1. No suplemento, o JavaScript chama uma nova API Office.js `getAccessTokenAsync`. Isso notifica o aplicativo host do Office para que obtenha um token de acesso para o suplemento. (Daqui por diante, ele ser? chamado de **token de acesso de inicializa??o** porque ser? substitu?do por um segundo token no processo posteriormente. Para ver um exemplo de token de acesso de inicializa??o decodificado, veja [Exemplo de token de acesso](sso-in-office-add-ins.md#example-access-token).)
1. Se o usu?rio n?o estiver conectado, o aplicativo host do Office abrir? uma janela pop-up para o usu?rio entrar.
1. Se essa ? a primeira vez que o usu?rio atual usa seu suplemento, ser? solicitado que ele d? o consentimento.
1. O aplicativo host do Office solicita o **token de acesso de inicializa??o** do ponto de extremidade v2.0 do Azure AD para o usu?rio atual.
1. O Azure AD envia o token de inicializa??o ao aplicativo host do Office.
1. O aplicativo host do Office envia o **token de acesso de inicializa??o** ao suplemento como parte do objeto de resultado que retornou pela chamada de `getAccessTokenAsync`.
1. O JavaScript no suplemento faz uma solicita??o HTTP a uma API da Web que est? hospedada no mesmo dom?nio totalmente qualificado que o suplemento e inclui o **token de acesso de inicializa??o** como prova de autoriza??o.  
1. O c?digo do servidor valida o **token de acesso de inicializa??o** de entrada.
1. O c?digo do servidor usa o fluxo ?em nome de? (definido em [Troca de Token do OAuth2](https://tools.ietf.org/html/draft-ietf-oauth-token-exchange-02) e o [daemon ou aplicativo para servidores para o cen?rio da API da Web do Azure](https://docs.microsoft.com/en-us/azure/active-directory/develop/active-directory-authentication-scenarios#daemon-or-server-application-to-web-api)) para obter um token de acesso para o Microsoft Graph em troca do token de acesso de inicializa??o.
1. O Azure AD retorna o token de acesso ao Microsoft Graph para o suplemento (e um token de atualiza??o, se o suplemento solicitar a permiss?o *offline_access*).
1. O c?digo do servidor armazena em cache o token de acesso ao Microsoft Graph.
1. O c?digo do servidor faz solicita??es para o Microsoft Graph e inclui o token de acesso ao Microsoft Graph.
1. O Microsoft Graph retorna os dados para o suplemento, que pode transmiti-los ? interface do usu?rio do suplemento.
1. Quando o token de acesso ao Microsoft Graph expira, o c?digo do servidor pode usar seu token de atualiza??o para obter um novo token de acesso ao Microsoft Graph.

## <a name="develop-an-sso-add-in-that-accesses-microsoft-graph"></a>Desenvolver um suplemento SSO que acesse o Microsoft Graph

Voc? desenvolve um suplemento que acessa o Microsoft Graph como faria com qualquer outro suplemento que usa SSO. Para uma descri??o detalhada, veja [Ativar logon ?nico para Suplementos do Office](https://docs.microsoft.com/en-us/office/dev/add-ins/develop/sso-in-office-add-ins). A diferen?a ? que ? obrigat?rio que o suplemento tenha uma API da Web do servidor e o que ? chamado de token de acesso nesse artigo ? chamado de "token de acesso de inicializa??o". 

Dependendo do seu idioma e estrutura, bibliotecas podem estar dispon?veis para simplificar o c?digo do servidor que voc? precisa escrever. Seu c?digo deve fazer o seguinte:

* Validar o token de acesso de inicializa??o do suplemento recebido do manipulador de token criado anteriormente. Para mais informa??es, veja [Validar o token de acesso](sso-in-office-add-ins.md#validate-the-access-token). 
* Iniciar o fluxo "em nome de" com uma chamada para o ponto de extremidade v2.0 do Azure AD que inclui o token de acesso de inicializa??o, alguns metadados sobre o usu?rio e as credenciais do suplemento (sua ID e segredo).
* Armazenar em cache o token de acesso retornado para o Microsoft Graph. Para mais informa??es sobre esse fluxo, veja [Fluxo em nome de do Active Directory do Azure v2.0 e OAuth 2.0](https://docs.microsoft.com/en-us/azure/active-directory/develop/active-directory-v2-protocols-oauth-on-behalf-of).
* Criar um ou mais m?todos de API da Web que obtenham dados do Microsoft Graph passando o token de acesso em cache para o Microsoft Graph.

> [!NOTE]
> Para ver exemplos de tokens de acesso decodificados para o Microsoft Graph e obtidos pelo fluxo "em nome de", veja [Fluxo em nome de do Active Directory do Azure v2.0 e OAuth 2.0](https://docs.microsoft.com/en-us/azure/active-directory/develop/active-directory-v2-protocols-oauth-on-behalf-of).

Confira exemplos de explica??es e cen?rios detalhados em:

* [Criar um Suplemento do Office com Node.js que usa logon ?nico](create-sso-office-add-ins-nodejs.md)
* [Criar um Suplemento do Office com ASP.NET que usa logon ?nico](create-sso-office-add-ins-aspnet.md)
* [Cen?rio: implementar o logon ?nico no servi?o em um suplemento do Outlook](https://docs.microsoft.com/en-us/outlook/add-ins/implement-sso-in-outlook-add-in)



