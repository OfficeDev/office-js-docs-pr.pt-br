---
title: Registrar um Suplemento do Office que usa SSO com o ponto de extremidade do Azure AD v2.0
description: Saiba como registrar um suplemento do Office com o ponto de extremidade do Azure AD v 2.0.
ms.date: 04/10/2019
localization_priority: Normal
ms.openlocfilehash: 8bcd72bd6f2d56c5f97d2d4f153d6791d111452e
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 06/08/2020
ms.locfileid: "44609370"
---
# <a name="register-an-office-add-in-that-uses-sso-with-the-azure-ad-v20-endpoint"></a>Registrar um Suplemento do Office que usa SSO com o ponto de extremidade do Azure AD v2.0

Este artigo explica como registrar um Suplemento do Office com o ponto de extremidade do Azure AD v2.0. É preciso registrar o suplemento ao iniciar o desenvolvimento. Ao progredir para teste ou produção, é possível alterar o registro existente ou criar registros separados para versões de desenvolvimento, teste e produção do suplemento.

A tabela a seguir relaciona as informações necessárias para executar este procedimento e os espaços reservados correspondentes que aparecem nas instruções.

|Informações  |Exemplos  |Espaço reservado  |
|---------|---------|---------|
|Um nome legível por humanos para o suplemento. (Recomenda-se exclusividade, mas não é obrigatória.)|`Contoso Marketing Excel Add-in (Prod)`|**$ADD-IN-NAME$**|
|O nome de domínio totalmente qualificado do suplemento (exceto para o protocolo). *Use um domínio pertencente a você.* Por esse motivo, você não pode usar determinados domínios conhecidos, como `azurewebsites.net` ou `cloudapp.net`. O domínio deve ser o mesmo, incluindo qualquer subdomínio, que o usado nas URLs na seção `<Resources>` do manifesto do suplemento.|`localhost:6789`, `addins.contoso.com`|**$FQDN-WITHOUT-PROTOCOL$**|
|As permissões para o AAD e o Microsoft Graph que seu suplemento precisa. (`profile` é sempre obrigatório.)|`profile`, `Files.Read.All`|N/D|

[!INCLUDE[](../includes/register-sso-add-in-aad-v2-include.md)]
