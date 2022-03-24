---
title: Autorizar para o Microsoft Graph de um Office Add-in
description: Aprenda a autorizar o microsoft Graph a partir de um Office Add-in.
ms.date: 01/25/2022
ms.localizationpriority: medium
ms.openlocfilehash: 8166b7a71767abd0456662dbe8573f59bb2c7e82
ms.sourcegitcommit: 968d637defe816449a797aefd930872229214898
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/23/2022
ms.locfileid: "63743580"
---
# <a name="authorize-to-microsoft-graph-from-an-office-add-in"></a>Autorizar para o Microsoft Graph de um Office Add-in

Seu add-in pode obter autorização para a Microsoft Graph dados obtendo um token de acesso para o Microsoft Graph do plataforma de identidade da Microsoft. Use o fluxo de Código de Autorização ou o fluxo implícito como faria em outros aplicativos Web, mas com uma exceção: o plataforma de identidade da Microsoft não permite que sua página de entrada seja aberta em um iframe. Quando um Suplemento do Office está sendo executado no *Office na Web*, o painel de tarefas é um iframe. Isso significa que você precisará abrir a página de login em uma caixa de diálogo usando a API de Office de diálogo. Isso afeta a maneira como você usa as bibliotecas auxiliares de autenticação e autorização. Para saber mais, confira [Autenticação com a API de Diálogo do Office](auth-with-office-dialog-api.md).

> [!NOTE]
> Se você estiver implementando o SSO e planeja acessar o Microsoft Graph, consulte Autorizar para a [Microsoft Graph com SSO](authorize-to-microsoft-graph.md).

Para obter informações sobre a autenticação de programação usando o plataforma de identidade da Microsoft, [consulte plataforma de identidade da Microsoft documentação](/azure/active-directory/develop). Você encontrará tutoriais e guias nesse conjunto de documentação, bem como links para exemplos relevantes. Mais uma vez, talvez seja necessário ajustar o código nos exemplos a serem executados na caixa de diálogo Office para levar em conta a caixa de diálogo Office que é executado em um processo separado do painel de tarefas.

Depois que seu código obtém o token de acesso para o Microsoft Graph, ele passa o token de acesso da caixa de diálogo para o painel de tarefas ou armazena o token em um banco de dados e sinaliza o painel de tarefas de que o token está disponível. (Consulte [Autenticação com a OFFICE de diálogo para](auth-with-office-dialog-api.md) obter detalhes.) O código no painel de tarefas solicita dados do Microsoft Graph e inclui o token nessas solicitações. Para obter mais informações sobre como chamar o Microsoft Graph e os SDKs do Microsoft Graph, consulte [Microsoft Graph documentação](/graph/).

## <a name="recommended-libraries-and-samples"></a>Bibliotecas e exemplos recomendados

Recomendamos que você use as seguintes bibliotecas ao acessar o Microsoft Graph.

- Para suplementos usando um lado do servidor com uma Estrutura baseada em rede, como o .NET Core ou o ASP.NET, use o[MSAL.NET](https://github.com/AzureAD/microsoft-authentication-library-for-dotnet/wiki#conceptual-documentation).
- Para suplementos usando um servidor baseado em NodeJS, use o[Passaport Azure AD.](https://github.com/AzureAD/passport-azure-ad)
- Para suplementos usando o fluxo implícito, use [MSAL. js.](https://github.com/AzureAD/microsoft-authentication-library-for-js/wiki)

Para obter mais informações sobre as bibliotecas recomendadas para trabalhar com a plataforma de identidade da Microsoft (o antigo AAD v.2.0), confira[bibliotecas de autenticação da plataforma de identidade da Microsoft](/azure/active-directory/develop/reference-v2-libraries).

Os exemplos a seguir Graph dados da Microsoft de um Office Add-in.

- [Suplemento do Office Microsoft Graph ASP.NET](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth/Office-Add-in-Microsoft-Graph-ASPNET)
- [Suplemento do Outlook Microsoft Graph ASP.NET](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth/Outlook-Add-in-Microsoft-Graph-ASPNET)
- [Suplemento do Office Microsoft Graph React](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth/Office-Add-in-Microsoft-Graph-React)
