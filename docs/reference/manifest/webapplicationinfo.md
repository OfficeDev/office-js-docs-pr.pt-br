# <a name="webapplicationinfo-element"></a><span data-ttu-id="2ca5c-101">Elemento WebApplicationInfo</span><span class="sxs-lookup"><span data-stu-id="2ca5c-101">WebApplicationInfo element</span></span>

<span data-ttu-id="2ca5c-102">Suporta o logon único (SSO) em Suplementos do Office. Este elemento contém informações sobre o suplemento como:</span><span class="sxs-lookup"><span data-stu-id="2ca5c-102">Supports single sign-on (SSO) in Office Add-ins. This element contains information for the add-in as both:</span></span>

- <span data-ttu-id="2ca5c-103">Um *recurso* do OAuth 2.0 para o qual o aplicativo de hospedagem do Office pode precisar de permissões.</span><span class="sxs-lookup"><span data-stu-id="2ca5c-103">An OAuth 2.0 *resource* to which the Office host application might need permissions.</span></span>
- <span data-ttu-id="2ca5c-104">Um *cliente* do OAuth 2.0 que pode exigir permissões para o Microsoft Graph.</span><span class="sxs-lookup"><span data-stu-id="2ca5c-104">An OAuth 2.0 *client* that might need permissions to Microsoft Graph.</span></span>

> [!NOTE]
> <span data-ttu-id="2ca5c-105">Atualmente a API de logon único tem suporte na visualização para Word, Excel, Outlook e PowerPoint.</span><span class="sxs-lookup"><span data-stu-id="2ca5c-105">The Single Sign-on API is currently supported in preview for Word, Excel, Outlook, and PowerPoint.</span></span> <span data-ttu-id="2ca5c-106">Saiba mais sobre os programas para os quais a API de logon único tem suporte no momento, consulte�[Conjuntos de requisitos da IdentityAPI](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/identity-api-requirement-sets?view=office-js).</span><span class="sxs-lookup"><span data-stu-id="2ca5c-106">For more information about where the Single Sign-on API is currently supported, see [IdentityAPI requirement sets](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/identity-api-requirement-sets?view=office-js).</span></span> <span data-ttu-id="2ca5c-107">Se você estiver trabalhando com um suplemento do Outlook, certifique-se de habilitar a Autenticação Moderna para o locatário do Office 365.</span><span class="sxs-lookup"><span data-stu-id="2ca5c-107">If you are working with an Outlook add-in, be sure to enable Modern Authentication for the Office 365 tenancy.</span></span> <span data-ttu-id="2ca5c-108">Para saber como fazer isso, consulte�[Exchange Online: como habilitar seu locatário para autenticação moderna](https://social.technet.microsoft.com/wiki/contents/articles/32711.exchange-online-how-to-enable-your-tenant-for-modern-authentication.aspx).</span><span class="sxs-lookup"><span data-stu-id="2ca5c-108">For information about how to do this, see [Exchange Online: How to enable your tenant for modern authentication](https://social.technet.microsoft.com/wiki/contents/articles/32711.exchange-online-how-to-enable-your-tenant-for-modern-authentication.aspx).</span></span>

<span data-ttu-id="2ca5c-109">**WebApplicationInfo** é um elemento filho do elemento [VersionOverrides](versionoverrides.md) no manifesto.</span><span class="sxs-lookup"><span data-stu-id="2ca5c-109">**WebApplicationInfo** is a child element of the [VersionOverrides](versionoverrides.md) element in the manifest.</span></span>  

## <a name="child-elements"></a><span data-ttu-id="2ca5c-110">Elementos filho</span><span class="sxs-lookup"><span data-stu-id="2ca5c-110">Child elements</span></span>

|  <span data-ttu-id="2ca5c-111">Elemento</span><span class="sxs-lookup"><span data-stu-id="2ca5c-111">Element</span></span> |  <span data-ttu-id="2ca5c-112">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="2ca5c-112">Required</span></span>  |  <span data-ttu-id="2ca5c-113">Descrição</span><span class="sxs-lookup"><span data-stu-id="2ca5c-113">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="2ca5c-114">**Id**</span><span class="sxs-lookup"><span data-stu-id="2ca5c-114">**Id**</span></span>    |  <span data-ttu-id="2ca5c-115">Sim</span><span class="sxs-lookup"><span data-stu-id="2ca5c-115">Yes</span></span>   |  <span data-ttu-id="2ca5c-116">A **Id do Aplicativo** do serviço associado do suplemento conforme registrado no ponto de extremidade do Azure Active Directory (Azure AD) v 2.0.</span><span class="sxs-lookup"><span data-stu-id="2ca5c-116">The **Application Id** of the add-in's associated service as registered in the Azure Active Directory v 2.0 endpoint.</span></span>|
|  <span data-ttu-id="2ca5c-117">**Recurso**</span><span class="sxs-lookup"><span data-stu-id="2ca5c-117">**Resource**</span></span>  |  <span data-ttu-id="2ca5c-118">Sim</span><span class="sxs-lookup"><span data-stu-id="2ca5c-118">Yes</span></span>   |  <span data-ttu-id="2ca5c-119">Especifica o **URI da ID do Aplicativo** do suplemento, conforme registrado no ponto de extremidade do Azure Active Directory v 2.0.</span><span class="sxs-lookup"><span data-stu-id="2ca5c-119">Specifies the **Application ID URI** of the add-in as registered in the Azure Active Directory v 2.0 endpoint.</span></span>|
|  [<span data-ttu-id="2ca5c-120">Escopos</span><span class="sxs-lookup"><span data-stu-id="2ca5c-120">Scopes</span></span>](scopes.md)                |  <span data-ttu-id="2ca5c-121">Não</span><span class="sxs-lookup"><span data-stu-id="2ca5c-121">No</span></span>  |  <span data-ttu-id="2ca5c-122">Especifica as permissões que seu suplemento precisa para o Microsoft Graph.</span><span class="sxs-lookup"><span data-stu-id="2ca5c-122">Specifies the permissions that the add-in needs to Microsoft Graph.</span></span>  |

> [!NOTE] 
> <span data-ttu-id="2ca5c-123">Atualmente, é necessário que o recurso do seu suplemento corresponda ao seu host.</span><span class="sxs-lookup"><span data-stu-id="2ca5c-123">Note: Currently, it's necessary that your add-in's Resource matches its Host.</span></span> <span data-ttu-id="2ca5c-124">O Office não solicitará um token para um suplemento, a menos que possa provar a propriedade, e hoje isso é feito hospedando o suplemento sob o nome de domínio totalmente qualificado do recurso.</span><span class="sxs-lookup"><span data-stu-id="2ca5c-124">Office will not request a Token for an add-in unless it can prove ownership, and today this is done by hosting the add-in under the Resource's fully-qualified domain name.</span></span>

## <a name="webapplicationinfo-example"></a><span data-ttu-id="2ca5c-125">Exemplo de WebApplicationInfo</span><span class="sxs-lookup"><span data-stu-id="2ca5c-125">WebApplicationInfo example</span></span>

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
