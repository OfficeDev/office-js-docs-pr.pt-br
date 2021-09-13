---
title: Elemento WebApplicationInfo no arquivo de manifesto
description: Documentação de referência do elemento WebApplicationInfo para Office arquivos XML (manifesto de complementos).
ms.date: 07/30/2020
ms.localizationpriority: medium
ms.openlocfilehash: 7de9271fc3e7ed76c0423c8a0b8ab70360b105c3
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 09/12/2021
ms.locfileid: "59152030"
---
# <a name="webapplicationinfo-element"></a>Elemento WebApplicationInfo

Suporta o logon único (SSO) em Suplementos do Office. Este elemento contém informações sobre o suplemento como:

- Um recurso OAuth 2.0  para o qual o Office cliente pode precisar de permissões.
- Um *cliente* do OAuth 2.0 que pode exigir permissões para o Microsoft Graph.

> [!NOTE]
> No momento, a API de login único tem suporte para Word, Excel, Outlook e PowerPoint. Para saber mais sobre os programas para os quais a API de logon único tem suporte no momento, consulte [Conjuntos de requisitos da IdentityAPI](../requirement-sets/identity-api-requirement-sets.md). Se você estiver trabalhando com um suplemento do Outlook, certifique-se de habilitar a Autenticação Moderna para a locação do Microsoft 365. Para saber como fazer isso, consulte [Exchange Online: como habilitar seu locatário para autenticação moderna](https://social.technet.microsoft.com/wiki/contents/articles/32711.exchange-online-how-to-enable-your-tenant-for-modern-authentication.aspx).

**WebApplicationInfo** é um elemento filho do elemento [VersionOverrides](versionoverrides.md) no manifesto.  

## <a name="child-elements"></a>Elementos filho

|  Elemento |  Obrigatório  |  Descrição  |
|:-----|:-----|:-----|
|  **Id**    |  Sim   |  A **Id do Aplicativo** do serviço associado do suplemento conforme registrado no ponto de extremidade do Azure Active Directory (Azure AD) v 2.0.|
|  **MsaId**    |  Não   |  A ID do cliente do aplicativo Web do seu complemento para MSA conforme registrado no msm.live.com.|
|  **Recurso**  |  Sim   |  Especifica o **URI da ID do Aplicativo** do suplemento, conforme registrado no ponto de extremidade do Azure Active Directory v 2.0.|
|  [Escopos](scopes.md)                |  Sim  |  Especifica as permissões que o complemento precisa para um recurso, como o Microsoft Graph.  |
|  [Autorizações](authorizations.md)  |  Não   | Especifica os recursos externos que o aplicativo Web do complemento precisa de autorização e as permissões necessárias.|

## <a name="webapplicationinfo-example"></a>Exemplo de WebApplicationInfo

```xml
<OfficeApp>
...
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides" xsi:type="VersionOverridesV1_0">
    ...
    <WebApplicationInfo>
      <Id>12345678-abcd-1234-efab-123456789abc</Id>
      <Resource>api://myDomain.com/12345678-abcd-1234-efab-123456789abc</Resource>
      <Scopes>
        <Scope>Files.Read.All</Scope>
        <Scope>offline_access</Scope>
        <Scope>openid</Scope>
        <Scope>profile</Scope>
      </Scopes>
      <Authorizations>
        <Authorization>
          <Resource>https://api.contoso.com</Resource>
            <Scopes>
              <Scope>profile</Scope>
          </Scopes>
        </Authorization>
      </Authorizations>
    </WebApplicationInfo>
  </VersionOverrides>
...
</OfficeApp>
```
