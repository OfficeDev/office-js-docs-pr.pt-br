---
title: Registrar um Suplemento do Office que usa SSO com o ponto de extremidade do Azure AD v2.0
description: ''
ms.date: 04/10/2018
ms.openlocfilehash: 95b690e21bddf7f2754cc308c8b771e629bbc630
ms.sourcegitcommit: c72c35e8389c47a795afbac1b2bcf98c8e216d82
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/23/2018
ms.locfileid: "19437252"
---
# <a name="register-an-office-add-in-that-uses-sso-with-the-azure-ad-v20-endpoint"></a>Registrar um Suplemento do Office que usa SSO com o ponto de extremidade do Azure AD v2.0

Este artigo explica como registrar um Suplemento do Office com o ponto de extremidade do Azure AD v2.0. É preciso registrar o suplemento ao começar a desenvolvê-lo. Ao progredir para o ambiente de teste ou produção, é possível alterar o registro existente ou criar registros separados para as versões de desenvolvimento, teste e produção do suplemento. 

A tabela a seguir relaciona as informações necessárias para executar este procedimento e os espaços reservados correspondentes que aparecem nas instruções. 

|Informação  |Exemplos  |Espaço reservado  |
|---------|---------|---------|
|Um nome legível para o suplemento. (Exclusividade recomendada, mas não obrigatória.)    |`Contoso Marketing Excel Add-in (Prod)`        |**$ADD-IN-NAME$**         |
|O nome de domínio totalmente qualificado (exceto para o protocolo) do suplemento. *É necessário usar um domínio possuído por você.* Por esse motivo, não é possível usar determinados domínios conhecidos, como `azurewebsites.net` ou `cloudapp.net`.   |`localhost:6789`, `addins.contoso.com`         |**$FQDN-WITHOUT-PROTOCOL$**         |
|As permissões para o AAD e o Microsoft Graph que o suplemento precisa. (`profile` é sempre necessário.)    |`profile`, `Files.Read.All`         |N/A         |

[!INCLUDE[](../includes/register-sso-add-in-aad-v2-include.md)]