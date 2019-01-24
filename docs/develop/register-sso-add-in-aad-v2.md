---
title: Registrar um Suplemento do Office que usa SSO com o ponto de extremidade do Azure AD v2.0
description: ''
ms.date: 04/10/2018
localization_priority: Priority
ms.openlocfilehash: c6cc775fc796f2235d8613a9ca9f6ee385721f85
ms.sourcegitcommit: d1aa7201820176ed986b9f00bb9c88e055906c77
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 01/23/2019
ms.locfileid: "29388175"
---
# <a name="register-an-office-add-in-that-uses-sso-with-the-azure-ad-v20-endpoint"></a>Registrar um Suplemento do Office que usa SSO com o ponto de extremidade do Azure AD v2.0

Este artigo explica como registrar um Suplemento do Office com o ponto de extremidade do Azure AD v2.0. É preciso registrar o suplemento ao iniciar o desenvolvimento. Ao progredir para teste ou produção, é possível alterar o registro existente ou criar registros separados para versões de desenvolvimento, teste e produção do suplemento.

A tabela a seguir relaciona as informações necessárias para executar este procedimento e os espaços reservados correspondentes que aparecem nas instruções. 

|Informações  |Exemplos  |Espaço reservado  |
|---------|---------|---------|
|Um nome legível por humanos para o suplemento. (Recomenda-se exclusividade, mas não é obrigatória.)    |`Contoso Marketing Excel Add-in (Prod)`        |**$ADD-IN-NAME$**         |
|O nome de domínio totalmente qualificado do suplemento (exceto para o protocolo). *Use um domínio pertencente a você.* Por esse motivo, você não pode usar determinados domínios conhecidos, como `azurewebsites.net` ou `cloudapp.net`. O domínio deve ser o mesmo, incluindo qualquer subdomínio, que o usado nas URLs na seção `<Resources>` do manifesto do suplemento.  |`localhost:6789`, `addins.contoso.com`         |**$FQDN-WITHOUT-PROTOCOL$**         |
|As permissões para o AAD e o Microsoft Graph que seu suplemento precisa. (`profile` é sempre obrigatório.)    |`profile`, `Files.Read.All`         |N/D         |

[!INCLUDE[](../includes/register-sso-add-in-aad-v2-include.md)]
