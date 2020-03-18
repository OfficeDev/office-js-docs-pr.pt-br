---
title: Elemento WebApplicationInfo no arquivo de manifesto
description: Documentação de referência do elemento VersionOverrides para arquivos de manifesto de suplementos do Office (XML).
ms.date: 08/12/2019
localization_priority: Normal
ms.openlocfilehash: 6acd0d5688bdd93d4054d0589afe5517afb1296f
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/17/2020
ms.locfileid: "42720292"
---
# <a name="webapplicationinfo-element"></a><span data-ttu-id="6ad26-103">Elemento WebApplicationInfo</span><span class="sxs-lookup"><span data-stu-id="6ad26-103">WebApplicationInfo element</span></span>

<span data-ttu-id="6ad26-104">Suporta o logon único (SSO) em Suplementos do Office. Este elemento contém informações sobre o suplemento como:</span><span class="sxs-lookup"><span data-stu-id="6ad26-104">Supports single sign-on (SSO) in Office Add-ins. This element contains information for the add-in as both:</span></span>

- <span data-ttu-id="6ad26-105">Um *recurso* do OAuth 2.0 para o qual o aplicativo de hospedagem do Office pode precisar de permissões.</span><span class="sxs-lookup"><span data-stu-id="6ad26-105">An OAuth 2.0 *resource* to which the Office host application might need permissions.</span></span>
- <span data-ttu-id="6ad26-106">Um *cliente* do OAuth 2.0 que pode exigir permissões para o Microsoft Graph.</span><span class="sxs-lookup"><span data-stu-id="6ad26-106">An OAuth 2.0 *client* that might need permissions to Microsoft Graph.</span></span>

> [!NOTE]
> <span data-ttu-id="6ad26-107">Atualmente, a API de logon único tem suporte para Word, Excel, Outlook e PowerPoint.</span><span class="sxs-lookup"><span data-stu-id="6ad26-107">The single sign-on API is currently supported in preview for Word, Excel, Outlook, and PowerPoint.</span></span> <span data-ttu-id="6ad26-108">Para saber mais sobre os programas para os quais a API de logon único tem suporte no momento, consulte [Conjuntos de requisitos da IdentityAPI](../requirement-sets/identity-api-requirement-sets.md).</span><span class="sxs-lookup"><span data-stu-id="6ad26-108">For more information about where the single sign-on API is currently supported, see [IdentityAPI requirement sets](../requirement-sets/identity-api-requirement-sets.md).</span></span> <span data-ttu-id="6ad26-109">Se você estiver trabalhando com um suplemento do Outlook, certifique-se de habilitar a Autenticação moderna para o locatário do Office 365.</span><span class="sxs-lookup"><span data-stu-id="6ad26-109">If you are working with an Outlook add-in, be sure to enable Modern Authentication for the Office 365 tenancy.</span></span> <span data-ttu-id="6ad26-110">Para saber como fazer isso, consulte [Exchange Online: como habilitar seu locatário para autenticação moderna](https://social.technet.microsoft.com/wiki/contents/articles/32711.exchange-online-how-to-enable-your-tenant-for-modern-authentication.aspx).</span><span class="sxs-lookup"><span data-stu-id="6ad26-110">To learn how to do this, see [Exchange Online: How to enable your tenant for modern authentication](https://social.technet.microsoft.com/wiki/contents/articles/32711.exchange-online-how-to-enable-your-tenant-for-modern-authentication.aspx).</span></span>

<span data-ttu-id="6ad26-111">**WebApplicationInfo** é um elemento filho do elemento [VersionOverrides](versionoverrides.md) no manifesto.</span><span class="sxs-lookup"><span data-stu-id="6ad26-111">**WebApplicationInfo** is a child element of the [VersionOverrides](versionoverrides.md) element in the manifest.</span></span>  

## <a name="child-elements"></a><span data-ttu-id="6ad26-112">Elementos filho</span><span class="sxs-lookup"><span data-stu-id="6ad26-112">Child elements</span></span>

|  <span data-ttu-id="6ad26-113">Elemento</span><span class="sxs-lookup"><span data-stu-id="6ad26-113">Element</span></span> |  <span data-ttu-id="6ad26-114">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="6ad26-114">Required</span></span>  |  <span data-ttu-id="6ad26-115">Descrição</span><span class="sxs-lookup"><span data-stu-id="6ad26-115">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="6ad26-116">**Id**</span><span class="sxs-lookup"><span data-stu-id="6ad26-116">**Id**</span></span>    |  <span data-ttu-id="6ad26-117">Sim</span><span class="sxs-lookup"><span data-stu-id="6ad26-117">Yes</span></span>   |  <span data-ttu-id="6ad26-118">A **Id do Aplicativo** do serviço associado do suplemento conforme registrado no ponto de extremidade do Azure Active Directory (Azure AD) v 2.0.</span><span class="sxs-lookup"><span data-stu-id="6ad26-118">The **Application Id** of the add-in's associated service as registered in the Azure Active Directory v 2.0 endpoint.</span></span>|
|  <span data-ttu-id="6ad26-119">**MsaId**</span><span class="sxs-lookup"><span data-stu-id="6ad26-119">**MsaId**</span></span>    |  <span data-ttu-id="6ad26-120">Não</span><span class="sxs-lookup"><span data-stu-id="6ad26-120">No</span></span>   |  <span data-ttu-id="6ad26-121">A ID do cliente do aplicativo Web do seu suplemento para o MSA, conforme registrado no msm.live.com.</span><span class="sxs-lookup"><span data-stu-id="6ad26-121">The client ID of your add-in's web application for MSA as registered in msm.live.com.</span></span>|
|  <span data-ttu-id="6ad26-122">**Recurso**</span><span class="sxs-lookup"><span data-stu-id="6ad26-122">**Resource**</span></span>  |  <span data-ttu-id="6ad26-123">Sim</span><span class="sxs-lookup"><span data-stu-id="6ad26-123">Yes</span></span>   |  <span data-ttu-id="6ad26-124">Especifica o **URI da ID do Aplicativo** do suplemento, conforme registrado no ponto de extremidade do Azure Active Directory v 2.0.</span><span class="sxs-lookup"><span data-stu-id="6ad26-124">Specifies the **Application ID URI** of the add-in as registered in the Azure Active Directory v 2.0 endpoint.</span></span>|
|  [<span data-ttu-id="6ad26-125">Escopos</span><span class="sxs-lookup"><span data-stu-id="6ad26-125">Scopes</span></span>](scopes.md)                |  <span data-ttu-id="6ad26-126">Sim</span><span class="sxs-lookup"><span data-stu-id="6ad26-126">Yes</span></span>  |  <span data-ttu-id="6ad26-127">Especifica as permissões que o suplemento precisa para um recurso, como o Microsoft Graph.</span><span class="sxs-lookup"><span data-stu-id="6ad26-127">Specifies the permissions that the add-in needs to a resource, such as Microsoft Graph.</span></span>  |
|  [<span data-ttu-id="6ad26-128">Autorizações</span><span class="sxs-lookup"><span data-stu-id="6ad26-128">Authorizations</span></span>](authorizations.md)  |  <span data-ttu-id="6ad26-129">Não</span><span class="sxs-lookup"><span data-stu-id="6ad26-129">No</span></span>   | <span data-ttu-id="6ad26-130">Especifica os recursos externos que o aplicativo Web do suplemento precisa de autorização e as permissões necessárias.</span><span class="sxs-lookup"><span data-stu-id="6ad26-130">Specifies the external resources that the add-in's web application needs authorization to and the required permissions.</span></span>|

## <a name="webapplicationinfo-example"></a><span data-ttu-id="6ad26-131">Exemplo de WebApplicationInfo</span><span class="sxs-lookup"><span data-stu-id="6ad26-131">WebApplicationInfo example</span></span>

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
