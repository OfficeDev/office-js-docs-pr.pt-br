# <a name="scopes-element"></a><span data-ttu-id="1a671-101">Elemento Scopes</span><span class="sxs-lookup"><span data-stu-id="1a671-101">Scopes element</span></span>

<span data-ttu-id="1a671-102">Contém as permissões para o Microsoft Graph de que o suplemento precisa.</span><span class="sxs-lookup"><span data-stu-id="1a671-102">Contains permissions to Microsoft Graph that the add-in needs.</span></span> <span data-ttu-id="1a671-103">Este elemento é usado pela Office Store para criar uma caixa de diálogo de consentimento.</span><span class="sxs-lookup"><span data-stu-id="1a671-103">The Office Store uses the Scopes element to create a consent dialog box.</span></span> <span data-ttu-id="1a671-104">Quando os usuários instalam o suplemento a partir da Office Store, eles são solicitados a conceder ao suplemento as permissões especificas para os dados do Microsoft Graph do usuário.</span><span class="sxs-lookup"><span data-stu-id="1a671-104">When users install the add-in from the Store, they are prompted to grant the add-in the specified permissions to the user's Microsoft Graph data.</span></span>

## <a name="child-elements"></a><span data-ttu-id="1a671-105">Elementos filho</span><span class="sxs-lookup"><span data-stu-id="1a671-105">Child elements</span></span>

|  <span data-ttu-id="1a671-106">Elemento</span><span class="sxs-lookup"><span data-stu-id="1a671-106">Element</span></span> |  <span data-ttu-id="1a671-107">Tipo</span><span class="sxs-lookup"><span data-stu-id="1a671-107">Type</span></span>  |  <span data-ttu-id="1a671-108">Descrição</span><span class="sxs-lookup"><span data-stu-id="1a671-108">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="1a671-109">**Scope**</span><span class="sxs-lookup"><span data-stu-id="1a671-109">**Scope**</span></span>                |  <span data-ttu-id="1a671-110">sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="1a671-110">string</span></span>     |   <span data-ttu-id="1a671-111">O nome de uma permissão para o Microsoft Graph; por exemplo, Files.Read.All.</span><span class="sxs-lookup"><span data-stu-id="1a671-111">The name of a permission to Microsoft Graph; for example, Files.Read.All.</span></span> |

## <a name="example"></a><span data-ttu-id="1a671-112">Exemplo</span><span class="sxs-lookup"><span data-stu-id="1a671-112">Example</span></span>

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
