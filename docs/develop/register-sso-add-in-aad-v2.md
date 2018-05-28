---
title: Registrar um Suplemento do Office que usa SSO com o ponto de extremidade do Azure AD v2.0
description: ''
ms.date: 04/10/2018
ms.openlocfilehash: 95b690e21bddf7f2754cc308c8b771e629bbc630
ms.sourcegitcommit: c72c35e8389c47a795afbac1b2bcf98c8e216d82
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/23/2018
---
# <a name="register-an-office-add-in-that-uses-sso-with-the-azure-ad-v20-endpoint"></a>Registrar um Suplemento do Office que usa SSO com o ponto de extremidade do Azure AD v2.0

Este artigo explica como registrar um Suplemento do Office com o ponto de extremidade do Azure AD v2.0. ? preciso registrar o suplemento ao come?ar a desenvolv?-lo. Ao progredir para o ambiente de teste ou produ??o, ? poss?vel alterar o registro existente ou criar registros separados para as vers?es de desenvolvimento, teste e produ??o do suplemento. 

A tabela a seguir relaciona as informa??es necess?rias para executar este procedimento e os espa?os reservados correspondentes que aparecem nas instru??es. 

|Informa??o  |Exemplos  |Espa?o reservado  |
|---------|---------|---------|
|Um nome leg?vel para o suplemento. (Exclusividade recomendada, mas n?o obrigat?ria.)    |`Contoso Marketing Excel Add-in (Prod)`        |**$ADD-IN-NAME$**         |
|O nome de dom?nio totalmente qualificado (exceto para o protocolo) do suplemento. *? necess?rio usar um dom?nio possu?do por voc?.* Por esse motivo, n?o ? poss?vel usar determinados dom?nios conhecidos, como `azurewebsites.net` ou `cloudapp.net`.   |`localhost:6789`, `addins.contoso.com`         |**$FQDN-WITHOUT-PROTOCOL$**         |
|As permiss?es para o AAD e o Microsoft Graph que o suplemento precisa. (`profile` ? sempre necess?rio.)    |`profile`, `Files.Read.All`         |N/A         |

[!INCLUDE[](../includes/register-sso-add-in-aad-v2-include.md)]