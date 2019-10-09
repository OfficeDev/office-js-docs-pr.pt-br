---
title: Autorizar o Microsoft Graph sem SSO
description: ''
ms.date: 08/07/2019
localization_priority: Priority
ms.openlocfilehash: 1d696783003fc475f98d5d1d37f60348aacec5eb
ms.sourcegitcommit: f781d7cfd980cd866d6d1d00c5b9d16c8a4b7f9b
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 09/20/2019
ms.locfileid: "37053309"
---
# <a name="authorize-to-microsoft-graph-without-sso"></a>Autorizar o Microsoft Graph sem SSO

Você pode conseguir autorização para o seu suplemento acessar os dados do Microsoft Graph obtendo um token de acesso ao Graph a partir do Azure Active Directory (AAD). Você faz isso usando o fluxo de código de autorização ou fluxo implícito da mesma forma que faria em qualquer outro aplicativo Web com uma exceção: o AAD não permite que sua página de logon seja aberta em um iframe. Quando um suplemento do Office está sendo executado no *Office na Web*, o painel de tarefas é um iframe. Isso significa que você precisará abrir a tela de logon do AAD em uma caixa de diálogo aberta com a API de diálogo do Office. Isso afetará a maneira como você usa as bibliotecas auxiliares de autenticação e autorização. Para saber mais, confira [Autenticação com a API de diálogo do Office](auth-with-office-dialog-api.md).

Para obter informações sobre a autenticação de programação com o AAD, comece com [Visão geral da plataforma de Identidade da Microsoft (v 2.0)](/azure/active-directory/develop/v2-overview). Há muitos tutoriais e guias nesse conjunto de documentos, bem como links para exemplos relevantes. Uma vez mais para não esquecer: Talvez seja necessário ajustar o código nos exemplos para execução na caixa de diálogo do Office pois devemos levar em consideração o fato de que a caixa de diálogo é executada em um processo separado do painel de tarefas.

Após o seu código obter o token de acesso para o Microsoft Graph, ele passa o token de acesso da caixa de diálogo para o painel de tarefas, ou armazena o token em um banco de dados e sinaliza o painel de tarefas no qual o token está disponível. (Confira [autenticação com a API de caixa de diálogo do Office](auth-with-office-dialog-api.md) para obter mais detalhes). O código no painel de tarefas solicita dados do Microsoft Graph e inclui o token nestas solicitações. Para obter mais informações sobre como chamar o Microsoft Graph e os SDKs para Microsoft Graph, confira a[ documentação do Microsoft Graph](/graph/).

## <a name="recommended-libraries-and-samples"></a>Bibliotecas e exemplos recomendados

Recomendamos que você use as seguintes bibliotecas ao acessar o Microsoft Graph sem usar o SSO:

- Para suplementos usando um lado do servidor com uma Estrutura baseada em rede, como o .NET Core ou o ASP.NET, use o[MSAL.NET](https://github.com/AzureAD/microsoft-authentication-library-for-dotnet/wiki#conceptual-documentation).
- Para suplementos usando um servidor baseado em NodeJS, use o[Passaport Azure AD.](https://github.com/AzureAD/passport-azure-ad)
- Para suplementos usando o fluxo implícito, use [MSAL. js.](https://github.com/AzureAD/microsoft-authentication-library-for-js/wiki)

Para obter mais informações sobre as bibliotecas recomendadas para trabalhar com a plataforma de identidade da Microsoft (o antigo AAD v.2.0), confira[bibliotecas de autenticação da plataforma de identidade da Microsoft](/azure/active-directory/develop/reference-v2-libraries).

Os exemplos a seguir recebem dados do Microsoft Graph de um suplemento do Office:

- [Suplemento do Office Microsoft Graph ASP.NET](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Samples/auth/Office-Add-in-Microsoft-Graph-ASPNET)
- [Suplemento do Outlook Microsoft Graph ASP.NET](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Samples/auth/Outlook-Add-in-Microsoft-Graph-ASPNET)

