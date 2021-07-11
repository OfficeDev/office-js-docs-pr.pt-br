---
title: Autorizar o Microsoft Graph sem SSO
description: Saiba como autorizar o Microsoft Graph sem SSO
ms.date: 01/29/2020
localization_priority: Normal
ms.openlocfilehash: ec8279f4112936692f0859fee2a1cc766c0b4342
ms.sourcegitcommit: 883f71d395b19ccfc6874a0d5942a7016eb49e2c
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 07/09/2021
ms.locfileid: "53350040"
---
# <a name="authorize-to-microsoft-graph-without-sso"></a>Autorizar o Microsoft Graph sem SSO

O seu complemento pode obter autorização para os dados do Microsoft Graph obtendo um token de acesso para o Microsoft Graph do Azure Active Directory (Azure AD). Use o fluxo de Código de Autorização ou o fluxo Implícito como faria em outros aplicativos Web, mas com uma exceção: o Azure AD não permite que sua página de entrada seja aberta em um iframe. Quando um Suplemento do Office está sendo executado no *Office na Web*, o painel de tarefas é um iframe. Isso significa que você precisará abrir a tela de logon do Azure AD em uma caixa de diálogo aberta com a API de Office de diálogo. Isso afeta a maneira como você usa as bibliotecas auxiliares de autenticação e autorização. Para saber mais, confira [Autenticação com a API de Diálogo do Office](auth-with-office-dialog-api.md).

Para obter informações sobre a autenticação de programação com o Azure AD, comece com plataforma de identidade da Microsoft [(v2.0) visão](/azure/active-directory/develop/v2-overview)geral , onde você encontrará tutoriais e guias nesse conjunto de documentação, bem como links para exemplos relevantes. Novamente, talvez seja necessário ajustar o código nos exemplos para execução na caixa de diálogo do Office pois devemos levar em consideração o fato de que a caixa de diálogo do Office é executada em um processo separado do painel de tarefas.

Depois que seu código obtém o token de acesso para o Microsoft Graph, ele passa o token de acesso da caixa de diálogo para o painel de tarefas ou armazena o token em um banco de dados e sinaliza o painel de tarefas de que o token está disponível. (Consulte [Authentication with the Office dialog API](auth-with-office-dialog-api.md) for details.) O código no painel de tarefas solicita dados do Microsoft Graph e inclui o token nessas solicitações. Para obter mais informações sobre como chamar o Microsoft Graph e os SDKs do Microsoft Graph, consulte [Microsoft Graph documentação.](/graph/)

## <a name="recommended-libraries-and-samples"></a>Bibliotecas e exemplos recomendados

Recomendamos que você use as seguintes bibliotecas ao acessar o Microsoft Graph sem usar o SSO.

- Para suplementos usando um lado do servidor com uma Estrutura baseada em rede, como o .NET Core ou o ASP.NET, use o[MSAL.NET](https://github.com/AzureAD/microsoft-authentication-library-for-dotnet/wiki#conceptual-documentation).
- Para suplementos usando um servidor baseado em NodeJS, use o[Passaport Azure AD.](https://github.com/AzureAD/passport-azure-ad)
- Para suplementos usando o fluxo implícito, use [MSAL. js.](https://github.com/AzureAD/microsoft-authentication-library-for-js/wiki)

Para obter mais informações sobre as bibliotecas recomendadas para trabalhar com a plataforma de identidade da Microsoft (o antigo AAD v.2.0), confira[bibliotecas de autenticação da plataforma de identidade da Microsoft](/azure/active-directory/develop/reference-v2-libraries).

Os exemplos a seguir Graph dados da Microsoft de um Office Add-in.

- [Suplemento do Office Microsoft Graph ASP.NET](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Samples/auth/Office-Add-in-Microsoft-Graph-ASPNET)
- [Suplemento do Outlook Microsoft Graph ASP.NET](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Samples/auth/Outlook-Add-in-Microsoft-Graph-ASPNET)
- [Suplemento do Office Microsoft Graph React](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Samples/auth/Office-Add-in-Microsoft-Graph-React)
