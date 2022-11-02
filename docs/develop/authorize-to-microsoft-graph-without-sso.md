---
title: Autorizar no Microsoft Graph por meio de um suplemento do Office
description: Saiba como autorizar o Microsoft Graph em um suplemento do Office.
ms.date: 01/25/2022
ms.localizationpriority: medium
ms.openlocfilehash: 37dd4be3acb92dc7884972de923d94936fa870f4
ms.sourcegitcommit: 3abcf7046446e7b02679c79d9054843088312200
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 11/02/2022
ms.locfileid: "68810166"
---
# <a name="authorize-to-microsoft-graph-from-an-office-add-in"></a>Autorizar no Microsoft Graph por meio de um suplemento do Office

Seu suplemento pode obter autorização para dados do Microsoft Graph obtendo um token de acesso ao Microsoft Graph do plataforma de identidade da Microsoft. Use o fluxo do Código de Autorização ou o fluxo implícito como faria em outros aplicativos Web, mas com uma exceção: o plataforma de identidade da Microsoft não permite que sua página de entrada seja aberta em um iframe. Quando um Suplemento do Office está em execução no *Office na Web*, o painel de tarefas é um iframe. Isso significa que você precisará abrir a página de entrada em uma caixa de diálogo usando a API de caixa de diálogo do Office. Isso afeta a maneira como você usa as bibliotecas auxiliares de autenticação e autorização. Para saber mais, confira [Autenticação com a API de Diálogo do Office](auth-with-office-dialog-api.md).

> [!NOTE]
> Se você estiver implementando o SSO e planeja acessar o Microsoft Graph, consulte [Autorizar no Microsoft Graph com SSO](authorize-to-microsoft-graph.md).

Para obter informações sobre a autenticação de programação usando o plataforma de identidade da Microsoft, consulte [plataforma de identidade da Microsoft documentação](/azure/active-directory/develop). Você encontrará tutoriais e guias nesse conjunto de documentação, bem como links para amostras relevantes. Mais uma vez, talvez seja necessário ajustar o código nos exemplos a serem executados na caixa de diálogo do Office para contabilizar a caixa de diálogo do Office que é executada em um processo separado do painel de tarefas.

Depois que o código obtém o token de acesso ao Microsoft Graph, ele passa o token de acesso da caixa de diálogo para o painel de tarefas ou armazena o token em um banco de dados e sinaliza o painel de tarefas de que o token está disponível. (Consulte [Autenticação com a API de diálogo do Office](auth-with-office-dialog-api.md) para obter detalhes.) O código no painel de tarefas solicita dados do Microsoft Graph e inclui o token nessas solicitações. Para obter mais informações sobre como chamar o Microsoft Graph e os SDKs do Microsoft Graph, confira [Documentação do Microsoft Graph](/graph/).

## <a name="recommended-libraries-and-samples"></a>Bibliotecas e exemplos recomendados

Recomendamos que você use as seguintes bibliotecas ao acessar o Microsoft Graph.

- Para suplementos usando um lado do servidor com uma Estrutura baseada em rede, como o .NET Core ou o ASP.NET, use o[MSAL.NET](https://github.com/AzureAD/microsoft-authentication-library-for-dotnet/wiki#conceptual-documentation).
- Para suplementos usando um servidor baseado em NodeJS, use o[Passaport Azure AD.](https://github.com/AzureAD/passport-azure-ad)
- Para suplementos usando o fluxo implícito, use [MSAL. js.](https://github.com/AzureAD/microsoft-authentication-library-for-js/wiki)

Para obter mais informações sobre as bibliotecas recomendadas para trabalhar com a plataforma de identidade da Microsoft (o antigo AAD v.2.0), confira[bibliotecas de autenticação da plataforma de identidade da Microsoft](/azure/active-directory/develop/reference-v2-libraries).

Os exemplos a seguir obtêm dados do Microsoft Graph de um Suplemento do Office.

- [Suplemento do Office Microsoft Graph ASP.NET](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth/Office-Add-in-Microsoft-Graph-ASPNET)
- [Suplemento do Outlook Microsoft Graph ASP.NET](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth/Outlook-Add-in-Microsoft-Graph-ASPNET)
- [Suplemento do Office Microsoft Graph React](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth/Office-Add-in-Microsoft-Graph-React)
