---
title: Elemento WebApplicationInfo no arquivo de manifesto
description: Documentação de referência do elemento VersionOverrides para arquivos de manifesto de suplementos do Office (XML).
ms.date: 07/30/2020
localization_priority: Normal
ms.openlocfilehash: b64739f388720abb872fe2407f7e7d3031b5ec70
ms.sourcegitcommit: 8fdd7369bfd97a273e222a0404e337ba2b8807b0
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 08/05/2020
ms.locfileid: "46573116"
---
# <a name="webapplicationinfo-element"></a><span data-ttu-id="78615-103">Elemento WebApplicationInfo</span><span class="sxs-lookup"><span data-stu-id="78615-103">WebApplicationInfo element</span></span>

<span data-ttu-id="78615-104">Suporta o logon único (SSO) em Suplementos do Office. Este elemento contém informações sobre o suplemento como:</span><span class="sxs-lookup"><span data-stu-id="78615-104">Supports single sign-on (SSO) in Office Add-ins. This element contains information for the add-in as both:</span></span>

- <span data-ttu-id="78615-105">Um *recurso* do OAuth 2.0 para o qual o aplicativo de hospedagem do Office pode precisar de permissões.</span><span class="sxs-lookup"><span data-stu-id="78615-105">An OAuth 2.0 *resource* to which the Office host application might need permissions.</span></span>
- <span data-ttu-id="78615-106">Um *cliente* do OAuth 2.0 que pode exigir permissões para o Microsoft Graph.</span><span class="sxs-lookup"><span data-stu-id="78615-106">An OAuth 2.0 *client* that might need permissions to Microsoft Graph.</span></span>

> [!NOTE]
> <span data-ttu-id="78615-107">Atualmente, a API de logon único tem suporte para Word, Excel, Outlook e PowerPoint.</span><span class="sxs-lookup"><span data-stu-id="78615-107">The single sign-on API is currently supported for Word, Excel, Outlook, and PowerPoint.</span></span> <span data-ttu-id="78615-108">Para saber mais sobre os programas para os quais a API de logon único tem suporte no momento, consulte [Conjuntos de requisitos da IdentityAPI](/office/dev/add-ins/reference/requirement-sets/identity-api-requirement-sets).</span><span class="sxs-lookup"><span data-stu-id="78615-108">For more information about where the single sign-on API is currently supported, see [IdentityAPI requirement sets](/office/dev/add-ins/reference/requirement-sets/identity-api-requirement-sets).</span></span> <span data-ttu-id="78615-109">Se você estiver trabalhando com um suplemento do Outlook, certifique-se de habilitar a Autenticação moderna para o locatário do Office 365.</span><span class="sxs-lookup"><span data-stu-id="78615-109">If you are working with an Outlook add-in, be sure to enable Modern Authentication for the Office 365 tenancy.</span></span> <span data-ttu-id="78615-110">Para saber como fazer isso, consulte [Exchange Online: como habilitar seu locatário para autenticação moderna](https://social.technet.microsoft.com/wiki/contents/articles/32711.exchange-online-how-to-enable-your-tenant-for-modern-authentication.aspx).</span><span class="sxs-lookup"><span data-stu-id="78615-110">To learn how to do this, see [Exchange Online: How to enable your tenant for modern authentication](https://social.technet.microsoft.com/wiki/contents/articles/32711.exchange-online-how-to-enable-your-tenant-for-modern-authentication.aspx).</span></span>

<span data-ttu-id="78615-111">**WebApplicationInfo** é um elemento filho do elemento [VersionOverrides](versionoverrides.md) no manifesto.</span><span class="sxs-lookup"><span data-stu-id="78615-111">**WebApplicationInfo** is a child element of the [VersionOverrides](versionoverrides.md) element in the manifest.</span></span>  

## <a name="child-elements"></a><span data-ttu-id="78615-112">Elementos filho</span><span class="sxs-lookup"><span data-stu-id="78615-112">Child elements</span></span>

|  <span data-ttu-id="78615-113">Elemento</span><span class="sxs-lookup"><span data-stu-id="78615-113">Element</span></span> |  <span data-ttu-id="78615-114">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="78615-114">Required</span></span>  |  <span data-ttu-id="78615-115">Descrição</span><span class="sxs-lookup"><span data-stu-id="78615-115">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="78615-116">**Id**</span><span class="sxs-lookup"><span data-stu-id="78615-116">**Id**</span></span>    |  <span data-ttu-id="78615-117">Sim</span><span class="sxs-lookup"><span data-stu-id="78615-117">Yes</span></span>   |  <span data-ttu-id="78615-118">A **Id do Aplicativo** do serviço associado do suplemento conforme registrado no ponto de extremidade do Azure Active Directory (Azure AD) v 2.0.</span><span class="sxs-lookup"><span data-stu-id="78615-118">The **Application Id** of the add-in's associated service as registered in the Azure Active Directory v 2.0 endpoint.</span></span>|
|  <span data-ttu-id="78615-119">**MsaId**</span><span class="sxs-lookup"><span data-stu-id="78615-119">**MsaId**</span></span>    |  <span data-ttu-id="78615-120">Não</span><span class="sxs-lookup"><span data-stu-id="78615-120">No</span></span>   |  <span data-ttu-id="78615-121">A ID do cliente do aplicativo Web do seu suplemento para o MSA, conforme registrado no msm.live.com.</span><span class="sxs-lookup"><span data-stu-id="78615-121">The client ID of your add-in's web application for MSA as registered in msm.live.com.</span></span>|
|  <span data-ttu-id="78615-122">**Recurso**</span><span class="sxs-lookup"><span data-stu-id="78615-122">**Resource**</span></span>  |  <span data-ttu-id="78615-123">Sim</span><span class="sxs-lookup"><span data-stu-id="78615-123">Yes</span></span>   |  <span data-ttu-id="78615-124">Especifica o **URI da ID do Aplicativo** do suplemento, conforme registrado no ponto de extremidade do Azure Active Directory v 2.0.</span><span class="sxs-lookup"><span data-stu-id="78615-124">Specifies the **Application ID URI** of the add-in as registered in the Azure Active Directory v 2.0 endpoint.</span></span>|
|  [<span data-ttu-id="78615-125">Escopos</span><span class="sxs-lookup"><span data-stu-id="78615-125">Scopes</span></span>](scopes.md)                |  <span data-ttu-id="78615-126">Sim</span><span class="sxs-lookup"><span data-stu-id="78615-126">Yes</span></span>  |  <span data-ttu-id="78615-127">Especifica as permissões que o suplemento precisa para um recurso, como o Microsoft Graph.</span><span class="sxs-lookup"><span data-stu-id="78615-127">Specifies the permissions that the add-in needs to a resource, such as Microsoft Graph.</span></span>  |
|  [<span data-ttu-id="78615-128">Autorizações</span><span class="sxs-lookup"><span data-stu-id="78615-128">Authorizations</span></span>](authorizations.md)  |  <span data-ttu-id="78615-129">Não</span><span class="sxs-lookup"><span data-stu-id="78615-129">No</span></span>   | <span data-ttu-id="78615-130">Especifica os recursos externos que o aplicativo Web do suplemento precisa de autorização e as permissões necessárias.</span><span class="sxs-lookup"><span data-stu-id="78615-130">Specifies the external resources that the add-in's web application needs authorization to and the required permissions.</span></span>|

## <a name="webapplicationinfo-example"></a><span data-ttu-id="78615-131">Exemplo de WebApplicationInfo</span><span class="sxs-lookup"><span data-stu-id="78615-131">WebApplicationInfo example</span></span>

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
