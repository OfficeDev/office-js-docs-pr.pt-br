# <a name="appdomain-element"></a><span data-ttu-id="a822a-101">Elemento AppDomain</span><span class="sxs-lookup"><span data-stu-id="a822a-101">AppDomain element</span></span>

<span data-ttu-id="a822a-102">Especifica um domínio adicional que será usado para carregar páginas na janela do suplemento.</span><span class="sxs-lookup"><span data-stu-id="a822a-102">Specifies an additional domain that will be used to load pages in the add-in window.</span></span>

<span data-ttu-id="a822a-103">**Tipo de suplemento:** Conteúdo, Painel de tarefas, Email</span><span class="sxs-lookup"><span data-stu-id="a822a-103">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="a822a-104">Sintaxe</span><span class="sxs-lookup"><span data-stu-id="a822a-104">Syntax</span></span>

```XML
<AppDomain>string</AppDomain>
```

> [!IMPORTANT]
> <span data-ttu-id="a822a-105">O valor do elemento **AppDomain** deve incluir o protocolo (ex., `<AppDomain>https://myappdomain<AppDomain>`).</span><span class="sxs-lookup"><span data-stu-id="a822a-105">The value of the **AppDomain** element must include the protocol (e.g., `<AppDomain>https://myappdomain<AppDomain>`).</span></span>

## <a name="contained-in"></a><span data-ttu-id="a822a-106">Contido em</span><span class="sxs-lookup"><span data-stu-id="a822a-106">Contained in</span></span>

[<span data-ttu-id="a822a-107">AppDomains</span><span class="sxs-lookup"><span data-stu-id="a822a-107">AppDomains</span></span>](appdomains.md)

## <a name="remarks"></a><span data-ttu-id="a822a-108">Comentários</span><span class="sxs-lookup"><span data-stu-id="a822a-108">Remarks</span></span>

<span data-ttu-id="a822a-109">Os elementos **AppDomain** deve ser usado para especificar os domínios adicionais diferentes daqueles especificados no elemento [SourceLocation](sourcelocation.md).</span><span class="sxs-lookup"><span data-stu-id="a822a-109">The  AppDomains and **AppDomain** elements are used to specify any additional domains other than the one specified in the [SourceLocation element. For more information, see Office Add-ins XML manifest](sourcelocation.md).</span></span> <span data-ttu-id="a822a-110">Confira mais informações em [Manifesto XML de Suplementos do Office](/office/dev/add-ins/develop/add-in-manifests).</span><span class="sxs-lookup"><span data-stu-id="a822a-110">For more information, see [Office Add-ins XML manifest](/office/dev/add-ins/develop/add-in-manifests).</span></span>
