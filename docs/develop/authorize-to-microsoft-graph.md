---
title: Autorizar para o Microsoft Graph no seu Suplemento do Office
description: ''
ms.date: 04/10/2018
ms.openlocfilehash: 83a9dd0beda9cb17a4f404c32cbe08a1e07f277e
ms.sourcegitcommit: 30435939ab8b8504c3dbfc62fd29ec6b0f1a7d22
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 09/12/2018
ms.locfileid: "23944296"
---
# <a name="authorize-to-microsoft-graph-in-your-office-add-in-preview"></a>Autorizar para o Microsoft Graph no seu Suplemento do Office (visualização)

Os usuários entram no Office (nas plataformas online, celular e área de trabalho) usando uma conta pessoal da Microsoft ou uma conta de corporativa ou de estudante (Office 365). A melhor forma para um suplemento do Office obter acesso autorizado ao [Microsoft Graph](https://developer.microsoft.com/graph/docs) é usar as credenciais de logon do usuário no Office. Isso permite que eles acessem seus dados do Microsoft Graph sem precisar entrar uma segunda vez. 

> [!NOTE]
> Atualmente a API de logon único tem suporte em versão prévia para Word, Excel e PowerPoint. Confira mais informações sobre os programas para os quais a API de logon único tem suporte no momento em [Conjuntos de requisitos da IdentityAPI](https://docs.microsoft.com/javascript/office/requirement-sets/identity-api-requirement-sets?view=office-js).
> Se você estiver trabalhando com um suplemento do Outlook, certifique-se de habilitar a Autenticação Moderna para a locação do Office 365. Confira mais informações sobre como fazer isso em [Exchange Online: como habilitar seu locatário para autenticação moderna](https://social.technet.microsoft.com/wiki/contents/articles/32711.exchange-online-how-to-enable-your-tenant-for-modern-authentication.aspx).

## <a name="add-in-architecture-for-sso-and-microsoft-graph"></a>Arquitetura do suplemento para SSO e Microsoft Graph

Além de hospedar as páginas e o JavaScript do aplicativo web, o suplemento também deve hospedar, ao mesmo tempo o [nome de domínio totalmente qualificado](https://docs.microsoft.com/windows/desktop/DNS/f-gly#_dns_fully_qualified_domain_name_fqdn__gly), uma ou mais APIs web que obterá um token de acesso ao Microsoft Graph e fará solicitações a ele.

O manifesto do suplemento contém a marcação que especifica como ele está registrado no ponto de extremidade v2.0 do Azure Active Directory (Azure AD) e especifica todas as permissões para o Microsoft Graph que o suplemento precisa.

### <a name="how-it-works-at-runtime"></a>Como ele funciona em tempo de execução

O diagrama a seguir mostra como o processo de entrada e acesso ao Microsoft Graph funciona.

![Diagrama que mostra o processo de SSO](../images/sso-access-to-microsoft-graph.png)

1. No suplemento, o JavaScript chama a nova API Office.js [ getAccessTokenAsync](https://docs.microsoft.com/office/dev/add-ins/develop/sso-in-office-add-ins#sso-api-reference) . Isso informa ao aplicativo host do Office para obter um token de acesso ao suplemento. (A partir de agora, ele será chamado de ** token de acesso de inicialização**  porque será substituído por um segundo token mais tarde no processo. Para um exemplo de token de acesso de inicialização decodificado, confira [  Exemplo de token de acesso](sso-in-office-add-ins.md#example-access-token) .)
1. Se o usuário não estiver conectado, o aplicativo host do Office abrirá uma janela pop-up para o usuário entrar.
1. Se essa for a primeira vez que o usuário atual usa seu suplemento, será solicitado que ele dê o consentimento.
1. O aplicativo host do Office solicita o **token de acesso de inicialização** do ponto de extremidade do Azure AD v2.0 para o usuário atual.
1. O Azure AD envia o token de inicialização ao aplicativo host do Office.
1. O aplicativo host do Office envia o **token de acesso de inicialização** ao suplemento como parte do objeto de resultado que retornou pela chamada de `getAccessTokenAsync`.
1. O JavaScript no suplemento faz uma solicitação HTTP a uma API da Web que está hospedada no mesmo domínio totalmente qualificado que o suplemento e inclui o **token de acesso de inicialização** como prova de autorização.  
1. O código do servidor valida o **token de acesso de inicialização** de entrada.
1. O código do servidor usa o fluxo “em nome de” (definido em [Troca de Token do OAuth2](https://tools.ietf.org/html/draft-ietf-oauth-token-exchange-02) e o [daemon ou aplicativo para servidores para o cenário da API da Web do Azure](https://docs.microsoft.com/azure/active-directory/develop/active-directory-authentication-scenarios#daemon-or-server-application-to-web-api)) para obter um token de acesso para o Microsoft Graph em troca do token de acesso de inicialização.
1. O Azure AD retorna o token de acesso ao Microsoft Graph (e um token de atualização, se o suplemento solicitar a permissão *offline_access*) para o suplemento.
1. O código do servidor armazena em cache o token de acesso ao Microsoft Graph.
1. O código do lado do servidor faz solicitações ao Microsoft Graph e inclui o token de acesso ao Microsoft Graph.
1. O Microsoft Graph retorna os dados para o suplemento, que pode transmiti-los à interface do usuário do suplemento.
1. Quando o token de acesso ao Microsoft Graph expira, o código do servidor pode usar seu token de atualização para obter um novo token de acesso ao Microsoft Graph.

## <a name="develop-an-sso-add-in-that-accesses-microsoft-graph"></a>Desenvolver um suplemento SSO que acesse o Microsoft Graph

Você desenvolve um suplemento que acessa o Microsoft Graph como faria com qualquer outro suplemento que usa SSO. Para uma descrição detalhada, veja [Ativar logon único para Suplementos do Office](https://docs.microsoft.com/office/dev/add-ins/develop/sso-in-office-add-ins). A diferença é que é obrigatório que o suplemento tenha uma API da Web do servidor e o que é chamado de token de acesso nesse artigo é chamado de "token de acesso de inicialização". 

Dependendo do seu idioma e estrutura, podem existir bibliotecas disponíveis para simplificar o código do lado do servidor que você precisar escrever. Seu código deve fazer o seguinte:

* Validar o token de acesso de inicialização do suplemento recebido do manipulador de token que você criou. Para mais informações, veja [Validar o token de acesso](sso-in-office-add-ins.md#validate-the-access-token). 
* Iniciar o fluxo "em nome de" com uma chamada para o ponto de extremidade v2.0 do Azure AD que inclui o token de acesso de inicialização, alguns metadados sobre o usuário e as credenciais do suplemento (seu ID e segredo).
* Armazenar em cache o token de acesso ao Microsoft Graph que foi retornado. Para mais informações sobre esse fluxo, veja [Fluxo em nome de do Active Directory do Azure v2.0 e OAuth 2.0](https://docs.microsoft.com/azure/active-directory/develop/active-directory-v2-protocols-oauth-on-behalf-of).
* Criar um ou mais métodos de API da Web que obtenham dados do Microsoft Graph passando o token de acesso em cache para o Microsoft Graph.

> [!NOTE]
> Para ver exemplos de tokens de acesso decodificados para o Microsoft Graph e obtidos pelo fluxo "em nome de", veja [Fluxo em nome de do Active Directory do Azure v2.0 e OAuth 2.0](https://docs.microsoft.com/azure/active-directory/develop/active-directory-v2-protocols-oauth-on-behalf-of).

Confira exemplos de passo a passo e cenários detalhados em:

* [Criar um Suplemento do Office com Node.js que usa logon único](create-sso-office-add-ins-nodejs.md)
* [Criar um Suplemento do Office com ASP.NET que usa logon único](create-sso-office-add-ins-aspnet.md)
* [Cenário: implementar o logon único no serviço em um suplemento do Outlook](https://docs.microsoft.com/outlook/add-ins/implement-sso-in-outlook-add-in)



