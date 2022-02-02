---
title: Registrar um Office que usa o SSO com o plataforma de identidade da Microsoft
description: Saiba como registrar um Office com o plataforma de identidade da Microsoft para usar o SSO com o Word, Excel, PowerPoint e Outlook.
ms.date: 01/25/2022
ms.localizationpriority: medium
ms.openlocfilehash: b11ce5130e020b049038631b9ae1c0e62fdadeab
ms.sourcegitcommit: 57e15f0787c0460482e671d5e9407a801c17a215
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 02/02/2022
ms.locfileid: "62320240"
---
# <a name="register-an-office-add-in-that-uses-single-sign-on-sso-with-the-microsoft-identity-platform"></a>Registrar um Office de usuário que usa o SSO (sign-on único) com o plataforma de identidade da Microsoft

Este artigo explica como registrar um Office com o plataforma de identidade da Microsoft para que você possa usar o SSO. Registre o complemento quando você começar a desenvolve-lo para que, ao avançar para testes ou produção, você possa alterar o registro existente ou criar registros separados para versões de desenvolvimento, teste e produção do complemento.

A tabela a seguir relaciona as informações necessárias para executar este procedimento e os espaços reservados correspondentes que aparecem nas instruções.

|Informações  |Exemplos  |Espaço reservado  |
|---------|---------|---------|
|Um nome legível por humanos para o suplemento. (Recomenda-se exclusividade, mas não é obrigatória.)|`Contoso Marketing Excel Add-in (Prod)`|N/D|
|Uma ID de aplicativo que o Azure gera para você como parte do processo de registro.|`c6c1f32b-5e55-4997-881a-753cc1d563b7`|`<application-id>`|
|O nome de domínio totalmente qualificado do suplemento (exceto para o protocolo). *Use um domínio pertencente a você.* Por esse motivo, você não pode usar determinados domínios conhecidos, como `azurewebsites.net` ou `cloudapp.net`. O domínio deve ser o mesmo, incluindo qualquer subdomínio, que o usado nas URLs na seção `<Resources>` do manifesto do suplemento.|`localhost:6789`, `addins.contoso.com`|`<fully-qualified-domain-name>`|
|As permissões para o plataforma de identidade da Microsoft e a Microsoft Graph que seu complemento precisa. (`profile` é sempre obrigatório.)|`profile`, `Files.Read.All`|N/D|

[!INCLUDE[](../includes/register-sso-add-in-aad-v2-include.md)]
