---
title: Elemento WebApplicationInfo no arquivo de manifesto
description: ''
ms.date: 10/09/2018
ms.openlocfilehash: ae1f2ea43e795d8f4ac2f634785fc69b0ca7a3d0
ms.sourcegitcommit: 60fd8a3ac4a6d66cb9e075ce7e0cde3c888a5fe9
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 12/28/2018
ms.locfileid: "27457625"
---
# <a name="webapplicationinfo-element"></a><span data-ttu-id="45c26-102">Elemento WebApplicationInfo</span><span class="sxs-lookup"><span data-stu-id="45c26-102">WebApplicationInfo element</span></span>

<span data-ttu-id="45c26-103">Suporta o logon único (SSO) em Suplementos do Office. Este elemento contém informações sobre o suplemento como:</span><span class="sxs-lookup"><span data-stu-id="45c26-103">Supports single sign-on (SSO) in Office Add-ins. This element contains information for the add-in as both:</span></span>

- <span data-ttu-id="45c26-104">Um *recurso* do OAuth 2.0 para o qual o aplicativo de hospedagem do Office pode precisar de permissões.</span><span class="sxs-lookup"><span data-stu-id="45c26-104">An OAuth 2.0 *resource* to which the Office host application might need permissions.</span></span>
- <span data-ttu-id="45c26-105">Um *cliente* do OAuth 2.0 que pode exigir permissões para o Microsoft Graph.</span><span class="sxs-lookup"><span data-stu-id="45c26-105">An OAuth 2.0 *client* that might need permissions to Microsoft Graph.</span></span>

> [!NOTE]
> <span data-ttu-id="45c26-106">Atualmente, a API de logon único tem suporte para Word, Excel, Outlook e PowerPoint.</span><span class="sxs-lookup"><span data-stu-id="45c26-106">The single sign-on API is currently supported in preview for Word, Excel, Outlook, and PowerPoint.</span></span> <span data-ttu-id="45c26-107">Para saber mais sobre os programas para os quais a API de logon único tem suporte no momento, consulte [Conjuntos de requisitos da IdentityAPI](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/identity-api-requirement-sets).</span><span class="sxs-lookup"><span data-stu-id="45c26-107">For more information about where the Single Sign-on API is currently supported, see [IdentityAPI requirement sets](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/identity-api-requirement-sets).</span></span> <span data-ttu-id="45c26-108">Se você estiver trabalhando com um suplemento do Outlook, certifique-se de habilitar a Autenticação moderna para o locatário do Office 365.</span><span class="sxs-lookup"><span data-stu-id="45c26-108">If you are working with an Outlook add-in, be sure to enable Modern Authentication for the Office 365 tenancy.</span></span> <span data-ttu-id="45c26-109">Para saber como fazer isso, consulte [Exchange Online: como habilitar seu locatário para autenticação moderna](https://social.technet.microsoft.com/wiki/contents/articles/32711.exchange-online-how-to-enable-your-tenant-for-modern-authentication.aspx).</span><span class="sxs-lookup"><span data-stu-id="45c26-109">To learn how to do this, see�[Exchange Online: How to enable your tenant for modern authentication](https://social.technet.microsoft.com/wiki/contents/articles/32711.exchange-online-how-to-enable-your-tenant-for-modern-authentication.aspx).</span></span>

<span data-ttu-id="45c26-110">**WebApplicationInfo** é um elemento filho do elemento [VersionOverrides](versionoverrides.md) no manifesto.</span><span class="sxs-lookup"><span data-stu-id="45c26-110">**WebApplicationInfo** is a child element of the [VersionOverrides](versionoverrides.md) element in the manifest.</span></span>  

## <a name="child-elements"></a><span data-ttu-id="45c26-111">Elementos filho</span><span class="sxs-lookup"><span data-stu-id="45c26-111">Child elements</span></span>

|  <span data-ttu-id="45c26-112">Elemento</span><span class="sxs-lookup"><span data-stu-id="45c26-112">Element</span></span> |  <span data-ttu-id="45c26-113">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="45c26-113">Required</span></span>  |  <span data-ttu-id="45c26-114">Descrição</span><span class="sxs-lookup"><span data-stu-id="45c26-114">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="45c26-115">**Id**</span><span class="sxs-lookup"><span data-stu-id="45c26-115">**Id**</span></span>    |  <span data-ttu-id="45c26-116">Sim</span><span class="sxs-lookup"><span data-stu-id="45c26-116">Yes</span></span>   |  <span data-ttu-id="45c26-117">A **Id do Aplicativo** do serviço associado do suplemento conforme registrado no ponto de extremidade do Azure Active Directory (Azure AD) v 2.0.</span><span class="sxs-lookup"><span data-stu-id="45c26-117">The **Application Id** of the add-in's associated service as registered in the Azure Active Directory v 2.0 endpoint.</span></span>|
|  <span data-ttu-id="45c26-118">**Recurso**</span><span class="sxs-lookup"><span data-stu-id="45c26-118">**Resource**</span></span>  |  <span data-ttu-id="45c26-119">Sim</span><span class="sxs-lookup"><span data-stu-id="45c26-119">Yes</span></span>   |  <span data-ttu-id="45c26-120">Especifica o **URI da ID do Aplicativo** do suplemento, conforme registrado no ponto de extremidade do Azure Active Directory v 2.0.</span><span class="sxs-lookup"><span data-stu-id="45c26-120">Specifies the **Application ID URI** of the add-in as registered in the Azure Active Directory v 2.0 endpoint.</span></span>|
|  [<span data-ttu-id="45c26-121">Escopos</span><span class="sxs-lookup"><span data-stu-id="45c26-121">Scopes</span></span>](scopes.md)                |  <span data-ttu-id="45c26-122">Não</span><span class="sxs-lookup"><span data-stu-id="45c26-122">No</span></span>  |  <span data-ttu-id="45c26-123">Especifica as permissões que seu suplemento precisa para o Microsoft Graph.</span><span class="sxs-lookup"><span data-stu-id="45c26-123">Specifies the permissions that the add-in needs to Microsoft Graph.</span></span>  |

> [!NOTE] 
> <span data-ttu-id="45c26-124">Atualmente, é necessário que o recurso do seu suplemento corresponda ao seu host.</span><span class="sxs-lookup"><span data-stu-id="45c26-124">Currently, it's necessary that your add-in's Resource matches its Host.</span></span> <span data-ttu-id="45c26-125">O Office não solicitará um token para um suplemento, a menos que possa provar a propriedade, e hoje isso é feito hospedando o suplemento sob o nome de domínio totalmente qualificado do recurso.</span><span class="sxs-lookup"><span data-stu-id="45c26-125">Office will not request a Token for an add-in unless it can prove ownership, and today this is done by hosting the add-in under the Resource's fully-qualified domain name.</span></span>

## <a name="webapplicationinfo-example"></a><span data-ttu-id="45c26-126">Exemplo de WebApplicationInfo</span><span class="sxs-lookup"><span data-stu-id="45c26-126">WebApplicationInfo example</span></span>

```xml
<OfficeApp>
...
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides" xsi:type="VersionOverridesV1_0">
    ...
    <WebApplicationInfo>
      <Id>12345678-abcd-1234-efab-123456789abc</Id>
      <Resource>api://myDomain.com/12345678-abcd-1234-efab-123456789abc<Resource>
      <Scopes>
        <Scope>Files.Read.All</Scope>
        <Scope>offline_access</Scope>
        <Scope>openid</Scope>
        <Scope>profile</Scope>        
      </Scopes>
    </WebApplicationInfo>
  </VersionOverrides>
...
</OfficeApp>
```
