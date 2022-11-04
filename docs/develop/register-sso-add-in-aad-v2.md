---
title: Registrar um Suplemento do Office que usa o SSO com o plataforma de identidade da Microsoft
description: Saiba como registrar um Suplemento do Office com o plataforma de identidade da Microsoft usar o SSO com Word, Excel, PowerPoint e Outlook.
ms.date: 10/06/2022
ms.localizationpriority: medium
ms.openlocfilehash: 0aab7d421ac57d1436d68c659f5d820717bcb846
ms.sourcegitcommit: 693e9a9b24bb81288d41508cb89c02b7285c4b08
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 10/28/2022
ms.locfileid: "68842093"
---
# <a name="register-an-office-add-in-that-uses-single-sign-on-sso-with-the-microsoft-identity-platform"></a>Registrar um Suplemento do Office que usa o SSO (logon único) com o plataforma de identidade da Microsoft

Este artigo explica como registrar um Suplemento do Office com o plataforma de identidade da Microsoft para que você possa usar o SSO. Registre o suplemento quando começar a desenvolvê-lo para que, ao progredir para o teste ou produção, você possa alterar o registro existente ou criar registros separados para versões de desenvolvimento, teste e produção do suplemento.

A tabela a seguir relaciona as informações necessárias para executar este procedimento e os espaços reservados correspondentes que aparecem nas instruções.

|Informações  |Exemplos  |Espaço reservado  |
|---------|---------|---------|
|Um nome legível por humanos para o suplemento. (Recomenda-se exclusividade, mas não é obrigatória.)|`Contoso Marketing Excel Add-in (Prod)`|N/D|
|Uma ID do aplicativo que o Azure gera para você como parte do processo de registro.|`c6c1f32b-5e55-4997-881a-753cc1d563b7`|`<application-id>`|
|O nome de domínio totalmente qualificado do suplemento (exceto para o protocolo). *Use um domínio pertencente a você.* Por esse motivo, você não pode usar determinados domínios conhecidos, como `azurewebsites.net` ou `cloudapp.net`. O domínio deve ser o mesmo, incluindo quaisquer subdomínios, conforme é usado nas URLs na **\<Resources\>** seção do manifesto do suplemento.|`localhost:6789`, `addins.contoso.com`|`<fully-qualified-domain-name>`|
|As permissões para o plataforma de identidade da Microsoft e o Microsoft Graph que seu suplemento precisa. (`profile` é sempre obrigatório.)|`profile`, `Files.Read.All`|N/D|

[!INCLUDE [register-sso-add-in-aad-v2-include](../includes/register-sso-add-in-aad-v2-include.md)]